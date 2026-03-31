const puppeteer = require("puppeteer");
const readline = require("node:readline/promises");
const { stdin: input, stdout: output } = require("node:process");

const ORDERS_URL = "https://www.vinted.pl/my_orders";
const COMPLETED_FILTER = "[data-testid=\"my-orders-filter-completed\"]";
const ORDER_ITEM = "[data-testid=\"my-orders-item\"]";
const ORDER_TITLE = "[data-testid=\"my-orders-item--title\"]";
const ORDER_PRICE = "h3";
const LOCATION_SELECTOR = "span[aria-label^=\"Members location\"]";

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

const USER_DATA_DIR = process.env.VINTED_PROFILE_DIR || "";
const CHROME_PATH = process.env.VINTED_CHROME_PATH || "";
const WAIT_FOR_LOGIN = (process.env.VINTED_WAIT_FOR_LOGIN || "1") !== "0";
const HEADLESS = (process.env.VINTED_HEADLESS || "0") === "1";
const TARGETS_JSON = process.env.VINTED_TARGETS_JSON || "";
const MAX_SCROLLS = Number.parseInt(process.env.VINTED_MAX_SCROLLS || "12", 10);
const MAX_ITEMS = Number.parseInt(process.env.VINTED_MAX_ITEMS || "120", 10);
const MAX_STAGNANT = Number.parseInt(process.env.VINTED_MAX_STAGNANT || "3", 10);
const MAX_NO_MATCH = Number.parseInt(process.env.VINTED_MAX_NO_MATCH || "5", 10);
const MAX_DETAIL_CHECKS = Number.parseInt(process.env.VINTED_MAX_DETAIL_CHECKS || "50", 10);

function normalizeTitle(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[\s\u00A0]+/g, " ")
    .trim();
}

function parsePrice(value) {
  const cleaned = String(value || "").replace(/[^0-9,.-]/g, "").replace(" ", "");
  const normalized = cleaned.replace(",", ".");
  const num = Number.parseFloat(normalized);
  return Number.isFinite(num) ? Number(num.toFixed(2)) : null;
}

function buildKey(title, price) {
  if (!title || price === null) return "";
  return `${title}||${price.toFixed(2)}`;
}

function parseTargets(raw) {
  if (!raw) return [];
  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function buildTargetCounts(targets) {
  const counts = new Map();
  for (const target of targets) {
    const title = normalizeTitle(target.title || target.title_original || "");
    const price = parsePrice(target.amount || target.price || "");
    const key = buildKey(title, price);
    if (!key) continue;
    counts.set(key, (counts.get(key) || 0) + 1);
  }
  return counts;
}

function buildTitleCounts(targets) {
  const counts = new Map();
  for (const target of targets) {
    const title = normalizeTitle(target.title || target.title_original || "");
    if (!title) continue;
    counts.set(title, (counts.get(title) || 0) + 1);
  }
  return counts;
}

function buildPriceCounts(targets) {
  const counts = new Map();
  for (const target of targets) {
    const price = parsePrice(target.amount || target.price || "");
    if (price === null) continue;
    const key = price.toFixed(2);
    counts.set(key, (counts.get(key) || 0) + 1);
  }
  return counts;
}

function selectOrdersForTargets(orders, titleCounts, priceCounts) {
  const selected = [];
  const usedTitles = new Map();
  const usedBundles = new Map();
  const usedPrices = new Map();

  for (const order of orders) {
    const titleKey = normalizeTitle(order.title || "");
    const priceValue = parsePrice(order.price || "");
    const isBundle = titleKey.startsWith("zestaw");

    if (titleKey && titleCounts.has(titleKey)) {
      const needed = titleCounts.get(titleKey) || 0;
      const have = usedTitles.get(titleKey) || 0;
      if (have < needed) {
        selected.push(order);
        usedTitles.set(titleKey, have + 1);
        continue;
      }
    }

    if (isBundle && priceValue !== null) {
      const priceKey = priceValue.toFixed(2);
      const needed = priceCounts.get(priceKey) || 0;
      const have = usedBundles.get(priceKey) || 0;
      if (needed > 0 && have < needed) {
        selected.push(order);
        usedBundles.set(priceKey, have + 1);
        continue;
      }
    }

    if (priceValue !== null) {
      const priceKey = priceValue.toFixed(2);
      const needed = priceCounts.get(priceKey) || 0;
      const have = usedPrices.get(priceKey) || 0;
      if (needed > 0 && have < needed) {
        selected.push(order);
        usedPrices.set(priceKey, have + 1);
      }
    }
  }

  return selected;
}

async function collectOrders(page, targetCounts) {
  const seen = new Map();
  const matched = new Map();
  const titleMatched = new Map();
  let lastTitleMatchedTotal = 0;
  let titleNoMatchStreak = 0;
  let stagnant = 0;
  let noMatchStreak = 0;
  const targetTotal = Array.from(targetCounts.values()).reduce((a, b) => a + b, 0);

  const targets = parseTargets(TARGETS_JSON);
  const titleCounts = buildTitleCounts(targets);
  const priceCounts = buildPriceCounts(targets);
  const titleTargetTotal = Array.from(titleCounts.values()).reduce((a, b) => a + b, 0);

  if (targetTotal > 0) {
    console.log(`Cel dopasowan: ${targetTotal}`);
  }
  if (titleTargetTotal > 0) {
    console.log(`Cel dopasowan tytulow: ${titleTargetTotal}`);
  }

  for (let i = 0; i < MAX_SCROLLS; i += 1) {
    await page.waitForSelector(ORDER_ITEM, { timeout: 30000 });

    let items = [];
    try {
      items = await page.$$eval(ORDER_ITEM, (nodes) =>
        nodes.map((node) => {
          const link = node.getAttribute("href") || "";
          const titleEl = node.querySelector("[data-testid=\"my-orders-item--title\"]");
          const priceEl = Array.from(node.querySelectorAll("h3")).find((el) =>
            (el.textContent || "").includes("zł")
          );
          return {
            href: link,
            title: titleEl ? titleEl.textContent.trim() : "",
            price: priceEl ? priceEl.textContent.trim() : "",
          };
        })
      );
    } catch (err) {
      const msg = String(err && err.message ? err.message : err);
      if (msg.includes("detached Frame")) {
        await page.goto(ORDERS_URL, { waitUntil: "networkidle2" });
        await page.waitForSelector(COMPLETED_FILTER, { timeout: 30000 });
        await page.click(COMPLETED_FILTER);
        await sleep(1200);
        continue;
      }
      throw err;
    }

    const before = seen.size;
    for (const item of items) {
      if (!item.href) continue;
      if (!seen.has(item.href)) {
        seen.set(item.href, item);
      }

      if (targetTotal > 0) {
        const key = buildKey(normalizeTitle(item.title), parsePrice(item.price));
        if (key && targetCounts.has(key)) {
          matched.set(key, Math.min(targetCounts.get(key), (matched.get(key) || 0) + 1));
        }
      }

      if (titleTargetTotal > 0) {
        const titleKey = normalizeTitle(item.title);
        if (titleKey && titleCounts.has(titleKey)) {
          titleMatched.set(
            titleKey,
            Math.min(titleCounts.get(titleKey), (titleMatched.get(titleKey) || 0) + 1)
          );
        }
      }
    }

    if (seen.size === before) {
      stagnant += 1;
    } else {
      stagnant = 0;
    }

    const matchedTotal = Array.from(matched.values()).reduce((a, b) => a + b, 0);
    const titleMatchedTotal = Array.from(titleMatched.values()).reduce((a, b) => a + b, 0);
    console.log(
      `Scroll ${i + 1}/${MAX_SCROLLS}: widoczne=${seen.size}, dopasowane=${matchedTotal}/${targetTotal || 0}`
    );
    if (titleTargetTotal > 0 && titleMatchedTotal >= titleTargetTotal) {
      console.log("Znaleziono wszystkie tytuly z maili, koncze przewijanie.");
      break;
    }
    if (titleMatchedTotal === lastTitleMatchedTotal) {
      titleNoMatchStreak += 1;
    } else {
      titleNoMatchStreak = 0;
      lastTitleMatchedTotal = titleMatchedTotal;
    }
    if (titleTargetTotal > 0 && titleNoMatchStreak >= MAX_NO_MATCH) {
      console.log("Brak nowych tytulow przez kilka scrolli, koncze przewijanie.");
      break;
    }
    if (targetTotal > 0) {
      if (matchedTotal >= targetTotal) {
        console.log("Znaleziono wszystkie potrzebne pozycje, koncze przewijanie.");
        break;
      }
      if (matchedTotal === 0) {
        noMatchStreak += 1;
      } else {
        noMatchStreak = 0;
      }
      if (noMatchStreak >= MAX_NO_MATCH) {
        console.log("Brak dopasowan przez kilka scrolli, koncze przewijanie.");
        break;
      }
    }
    if (stagnant >= MAX_STAGNANT) {
      console.log("Brak nowych kart po przewinieciach, koncze przewijanie.");
      break;
    }
    if (seen.size >= MAX_ITEMS) {
      console.log("Osiagnieto limit kart, koncze przewijanie.");
      break;
    }

    await page.evaluate("window.scrollTo(0, document.body.scrollHeight)");
    await sleep(1200);
  }

  const allOrders = Array.from(seen.values());
  if (!titleCounts.size && !priceCounts.size) {
    return allOrders;
  }
  return selectOrdersForTargets(allOrders, titleCounts, priceCounts);
}

function extractCountry(label) {
  if (!label) return "";
  const trimmed = label.replace(/^Members location\s*/i, "").trim();
  const parts = trimmed.split(",");
  return (parts[0] || "").trim();
}

function extractTransactionNumber(text) {
  if (!text) return "";
  const match = text.match(/Numer\s+transakcji\s*[:#]?\s*#?\s*(\d{6,})/i);
  if (match) return match[1];
  const hash = text.match(/#(\d{6,})/);
  return hash ? hash[1] : "";
}

async function getLocationLabel(page) {
  for (let attempt = 0; attempt < 5; attempt += 1) {
    try {
      await page.waitForSelector(LOCATION_SELECTOR, { timeout: 1500 });
    } catch {
      // Ignore and try scrolling below.
    }

    const label = await page
      .$eval(LOCATION_SELECTOR, (el) => el.getAttribute("aria-label") || "")
      .catch(() => "");
    if (label) return label;

    await page.evaluate(() => {
      window.scrollTo(0, 0);
      const scrollables = Array.from(document.querySelectorAll("*")).filter(
        (el) => el.scrollHeight > el.clientHeight + 20
      );
      for (const el of scrollables) {
        el.scrollTop = 0;
      }
    });
    await sleep(600);
  }

  return "";
}

(async () => {
  const launchOptions = {
    headless: HEADLESS,
    defaultViewport: null,
    args: ["--start-maximized", "--disable-blink-features=AutomationControlled"],
  };
  if (USER_DATA_DIR) {
    launchOptions.userDataDir = USER_DATA_DIR;
  }
  if (CHROME_PATH) {
    launchOptions.executablePath = CHROME_PATH;
  }

  const browser = await puppeteer.launch(launchOptions);
  const page = await browser.newPage();

  await page.goto(ORDERS_URL, { waitUntil: "networkidle2" });
  if (WAIT_FOR_LOGIN) {
    const rl = readline.createInterface({ input, output });
    try {
      await rl.question(
        "Zaloguj sie do Vinted w otwartym oknie (najlepiej login/haslo Vinted, bez Google), potem wroc i wcisnij ENTER..."
      );
    } finally {
      rl.close();
    }
  }

  try {
    await page.waitForSelector(COMPLETED_FILTER, { timeout: 300000 });
  } catch (err) {
    throw new Error(
      "Nie widze zakladki 'Zakonczone'. Najpierw zaloguj sie do Vinted w tym oknie przegladarki. " +
        "Jesli logowanie przez Google jest blokowane, uzyj konta Vinted lub ustaw VINTED_PROFILE_DIR z juz zalogowanym profilem."
    );
  }
  await page.click(COMPLETED_FILTER);
  await sleep(1500);
  await page.waitForSelector(ORDER_ITEM, { timeout: 30000 });

  const targets = parseTargets(TARGETS_JSON);
  console.log(`Cele do uzupelnienia: ${targets.length}`);
  const targetCounts = buildTargetCounts(targets);
  const titleCounts = buildTitleCounts(targets);
  const priceCounts = buildPriceCounts(targets);
  const targetTx = new Set(
    targets
      .map((target) => String(target.transaction_number || "").trim())
      .filter((value) => value)
  );
  const targetTitles = new Set(
    targets
      .map((target) => normalizeTitle(target.title || target.title_original || ""))
      .filter((value) => value)
  );
  const targetPrices = new Set(
    targets
      .map((target) => parsePrice(target.amount || target.price || ""))
      .filter((value) => value !== null)
      .map((value) => value.toFixed(2))
  );
  const orders = await collectOrders(page, targetCounts);
  const selectedOrders = selectOrdersForTargets(orders, titleCounts, priceCounts);
  console.log(`Zebrane karty zamowien: ${orders.length}`);
  console.log(`Wybrane karty do sprawdzenia: ${selectedOrders.length}`);

  const results = [];
  const foundCounts = new Map();
  const foundTitleCounts = new Map();
  const foundPriceCounts = new Map();
  let detailChecks = 0;
  for (const order of selectedOrders) {
    if (!order.href) continue;
    if (detailChecks >= MAX_DETAIL_CHECKS) {
      console.log("Osiagnieto limit szczegolow, koncze.");
      break;
    }
    const orderKey = buildKey(normalizeTitle(order.title), parsePrice(order.price));
    const matchesKey = targetCounts.size > 0 && orderKey && targetCounts.has(orderKey);
    const normalizedTitle = normalizeTitle(order.title);
    const titleHit = targetTitles.has(normalizedTitle);
    const priceValue = parsePrice(order.price);
    const priceHit = priceValue !== null && targetPrices.has(priceValue.toFixed(2));
    const isBundle = normalizedTitle.startsWith("zestaw");

    const neededKey = matchesKey ? targetCounts.get(orderKey) || 0 : 0;
    const haveKey = matchesKey ? foundCounts.get(orderKey) || 0 : 0;
    const allowByKey = matchesKey && haveKey < neededKey;

    const neededTitle = titleHit ? titleCounts.get(normalizedTitle) || 0 : 0;
    const haveTitle = titleHit ? foundTitleCounts.get(normalizedTitle) || 0 : 0;
    const allowByTitle = titleHit && neededTitle > haveTitle;

    const priceKey = priceValue !== null ? priceValue.toFixed(2) : "";
    const neededPrice = priceKey ? priceCounts.get(priceKey) || 0 : 0;
    const havePrice = priceKey ? foundPriceCounts.get(priceKey) || 0 : 0;
    const allowByPrice = priceKey && neededPrice > havePrice;

    if (!allowByKey && !allowByTitle && !(allowByPrice && (isBundle || !titleHit))) {
      continue;
    }
    const detail = page;
    detailChecks += 1;
    console.log(`Sprawdzam szczegoly: ${detailChecks}/${MAX_DETAIL_CHECKS}`);
    const url = order.href.startsWith("http") ? order.href : `https://www.vinted.pl${order.href}`;
    await detail.goto(url, { waitUntil: "networkidle2" });
    await sleep(1200);

    const locationLabel = await getLocationLabel(detail);
    const pageText = await detail.evaluate(() => document.body.innerText || "");

    const txNumber = extractTransactionNumber(pageText);
    const matchesTx = txNumber && targetTx.has(txNumber);
    let matchType = "";
    if (matchesTx) {
      matchType = "tx";
    } else if (allowByKey) {
      matchType = "title_price";
    } else if (allowByTitle) {
      matchType = "title_only";
    } else if (allowByPrice && isBundle) {
      matchType = "bundle_price";
    } else if (allowByPrice) {
      matchType = "price_only";
    }

    if (!matchType) {
      continue;
    }

    results.push({
      title: order.title || "",
      price: order.price || "",
      country: extractCountry(locationLabel),
      transaction_number: txNumber,
      order_url: url,
      match_type: matchType,
    });

    if (matchesTx) {
      targetTx.delete(txNumber);
    }

    if (matchType === "title_price" && orderKey) {
      foundCounts.set(orderKey, (foundCounts.get(orderKey) || 0) + 1);
    }
    if (matchType === "title_only" || matchType === "title_price") {
      foundTitleCounts.set(normalizedTitle, (foundTitleCounts.get(normalizedTitle) || 0) + 1);
    }
    if (matchType === "price_only" || matchType === "bundle_price") {
      if (priceKey) {
        foundPriceCounts.set(priceKey, (foundPriceCounts.get(priceKey) || 0) + 1);
      }
    }

  }

  console.log(JSON.stringify(results));
  await browser.close();
})().catch((err) => {
  console.error(err);
  process.exit(1);
});
