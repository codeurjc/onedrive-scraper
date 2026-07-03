/**
 * Persistent browser session management.
 *
 * Uses a Playwright *persistent context* so cookies and tokens survive between
 * runs — the user only completes the interactive (MFA) login once. The visible
 * browser window is what lets the user satisfy multi-factor prompts.
 */

import { chromium, type BrowserContext, type Page } from "playwright";

export interface BrowserSessionOptions {
  /** Directory holding the persisted browser profile. */
  sessionDir: string;
  /** Origin of the site, e.g. `https://org-my.sharepoint.com`. */
  origin: string;
  /** Site URL to load when checking/establishing authentication. */
  siteUrl: string;
  /** Show the window. Must be `true` for interactive login. */
  headed: boolean;
  /** Browser locale. */
  locale: string;
  /** Default navigation/operation timeout (ms). */
  timeout: number;
  /** Max time (ms) to wait for interactive login to complete. */
  loginTimeout: number;
  /** Optional logger; defaults to `console.error`. */
  log?: (message: string) => void;
}

const LOGIN_HOST_HINTS = [
  "login.microsoftonline.com",
  "login.live.com",
  "login.windows.net",
  "adfs",
];

/** Manages the persistent Playwright context and interactive login. */
export class BrowserSession {
  private context?: BrowserContext;
  private _page?: Page;
  private readonly log: (message: string) => void;

  constructor(private readonly options: BrowserSessionOptions) {
    this.log = options.log ?? ((m) => console.error(m));
  }

  /** The active page. Throws if the session has not been launched. */
  get page(): Page {
    if (!this._page) {
      throw new Error("Browser session not launched. Call launch() first.");
    }
    return this._page;
  }

  /** Whether the session has an active page. */
  get isOpen(): boolean {
    return !!this._page;
  }

  /** Launch the persistent context and open a page. */
  async launch(): Promise<void> {
    if (this.context) return;
    this.context = await chromium.launchPersistentContext(this.options.sessionDir, {
      headless: !this.options.headed,
      locale: this.options.locale,
      acceptDownloads: true,
      // Headed: use the real (maximized) window so login is comfortable.
      // Headless: force a wide viewport, otherwise Chromium defaults to
      // 800x600 and OneDrive switches to a compact layout that breaks the
      // DOM selectors (name cell then also contains the size).
      viewport: this.options.headed ? null : { width: 1600, height: 1000 },
      args: this.options.headed ? ["--start-maximized"] : [],
    });
    this.context.setDefaultTimeout(this.options.timeout);
    this.context.setDefaultNavigationTimeout(this.options.timeout);
    this._page =
      this.context.pages()[0] ?? (await this.context.newPage());
  }

  /**
   * Ensure the page is on the site origin (REST calls must be same-origin).
   * Navigates to the site URL if not already there.
   */
  async ensureOnOrigin(): Promise<void> {
    const page = this.page;
    if (!page.url().startsWith(this.options.origin)) {
      await page.goto(this.options.siteUrl, { waitUntil: "domcontentloaded" });
    }
  }

  /** True when the current session can call the SharePoint REST API. */
  async isAuthenticated(): Promise<boolean> {
    const page = this.page;
    try {
      await this.ensureOnOrigin();
    } catch {
      return false;
    }
    if (this.looksLikeLoginPage(page.url())) return false;
    try {
      const url = `${this.options.siteUrl}/_api/web/currentuser?$select=LoginName`;
      const ok = await page.evaluate(async (u) => {
        try {
          const res = await fetch(u, {
            headers: { Accept: "application/json;odata=nometadata" },
            credentials: "include",
          });
          return res.ok;
        } catch {
          return false;
        }
      }, url);
      return ok;
    } catch {
      return false;
    }
  }

  private looksLikeLoginPage(url: string): boolean {
    return LOGIN_HOST_HINTS.some((h) => url.includes(h));
  }

  /**
   * Ensure the user is authenticated, driving an interactive login if needed.
   * Resolves once REST calls succeed, or throws on timeout.
   */
  async ensureAuthenticated(): Promise<void> {
    await this.launch();

    if (await this.isAuthenticated()) {
      this.log("[onedrive-scraper] Reusing persisted authenticated session.");
      return;
    }

    if (!this.options.headed) {
      throw new Error(
        "Not authenticated and running headless. Re-run with a visible browser (headed: true) to log in interactively."
      );
    }

    this.log(
      "[onedrive-scraper] Authentication required. A browser window is open — " +
        "please sign in (including any multi-factor step). Waiting…"
    );
    await this.page.goto(this.options.siteUrl, { waitUntil: "domcontentloaded" });

    const deadline = Date.now() + this.options.loginTimeout;
    while (Date.now() < deadline) {
      await this.page.waitForTimeout(2000);
      if (await this.isAuthenticated()) {
        this.log("[onedrive-scraper] Login successful. Session persisted.");
        return;
      }
    }
    throw new Error(
      `Interactive login did not complete within ${Math.round(
        this.options.loginTimeout / 1000
      )}s.`
    );
  }

  /** Close the context, flushing the persisted profile to disk. */
  async close(): Promise<void> {
    await this.context?.close();
    this.context = undefined;
    this._page = undefined;
  }
}
