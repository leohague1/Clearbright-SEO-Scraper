const NOISE_DOMAINS = [
  "google.com", "googleapis.com", "example.com", "sentry.io", "w3.org",
  "schema.org", "placeholder.com", "yourdomain.com", "domain.com",
  "email.com", "test.com", "sample.com", "duckduckgo.com",
];

const EMAIL_REGEX = /[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/g;

export async function findEmail(page, businessName, town) {
  try {
    const query = `"${businessName}" "${town}" email`;
    // DuckDuckGo plain HTML — no JS, no CAPTCHA, scraper-friendly
    const url = `https://html.duckduckgo.com/html/?q=${encodeURIComponent(query)}&kl=uk-en`;

    await page.goto(url, { waitUntil: "domcontentloaded", timeout: 15000 });

    // Wait for results container — DDG HTML page loads synchronously so this is fast
    await page.waitForSelector(".results", { timeout: 8000 }).catch(() => {});

    const text = await page.evaluate(() => document.body.innerText);
    const matches = text.match(EMAIL_REGEX) || [];

    for (const email of matches) {
      const lower = email.toLowerCase();
      const localPart = lower.split("@")[0];
      const domain = lower.split("@")[1] || "";

      if (localPart.length > 40) continue;
      if (NOISE_DOMAINS.some((d) => domain === d || domain.endsWith(`.${d}`))) continue;

      return email;
    }

    return null;
  } catch {
    return null;
  }
}
