const puppeteer = require("puppeteer");
const readline = require("node:readline/promises");
const { stdin: input, stdout: output } = require("node:process");
const path = require("node:path");

const ORDERS_URL = "https://www.vinted.pl/my_orders";
const WALLET_HISTORY_BASE_URL = "https://www.vinted.pl/wallet/history";
const COMPLETED_FILTER = "[data-testid=\"my-orders-filter-completed\"]";
const ORDER_ITEM = "[data-testid=\"my-orders-item\"]";
const ORDER_TITLE = "[data-testid=\"my-orders-item--title\"]";
const ORDER_PRICE = "h3";
const LOCATION_SELECTOR = "span[aria-label^=\"Members location\"]";

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

function getArgValue(name) {
  const idx = process.argv.indexOf(name);
  if (idx >= 0 && idx + 1 < process.argv.length) {
    return String(process.argv[idx + 1] || "").trim();
  }
  return "";
}

function currentMonthString() {
  const now = new Date();
  const yyyy = String(now.getFullYear());
  const mm = String(now.getMonth() + 1).padStart(2, "0");
  return `${yyyy}-${mm}`;
}

const CLI_MODE = getArgValue("--mode") || (process.argv.includes("--history") ? "history" : "");
const CLI_MONTH = getArgValue("--month");
const CLI_ORDER_TYPES = getArgValue("--order-types");
const CLI_ITEM_QUERY = getArgValue("--item");

const DEFAULT_USER_DATA_DIR = path.join(__dirname, ".vinted-browser-profile");
const USER_DATA_DIR = process.env.VINTED_PROFILE_DIR || DEFAULT_USER_DATA_DIR;
const PROFILE_NAME = (process.env.VINTED_PROFILE_NAME || "").trim();
const CHROME_PATH = process.env.VINTED_CHROME_PATH || "";
const WAIT_FOR_LOGIN = (process.env.VINTED_WAIT_FOR_LOGIN || "1") !== "0";
const HEADLESS = (process.env.VINTED_HEADLESS || "0") === "1";
const TARGETS_JSON = process.env.VINTED_TARGETS_JSON || "";
const SCRAPE_MODE = String(CLI_MODE || process.env.VINTED_SCRAPE_MODE || "targets").trim().toLowerCase();
const HISTORY_TARGET_MONTH = String(CLI_MONTH || process.env.VINTED_SCRAPE_TARGET_MONTH || "").trim();
const HISTORY_ORDER_TYPES = String(CLI_ORDER_TYPES || process.env.VINTED_SCRAPE_ORDER_TYPES || "purchased,sold")
  .split(",")
  .map((v) => v.trim().toLowerCase())
  .filter((v) => v === "purchased" || v === "sold");
const MAX_SCROLLS = Number.parseInt(process.env.VINTED_MAX_SCROLLS || "12", 10);
const MAX_ITEMS = Number.parseInt(process.env.VINTED_MAX_ITEMS || "120", 10);
const MAX_STAGNANT = Number.parseInt(process.env.VINTED_MAX_STAGNANT || "3", 10);
const MAX_NO_MATCH = Number.parseInt(process.env.VINTED_MAX_NO_MATCH || "5", 10);
const MAX_DETAIL_CHECKS = Number.parseInt(process.env.VINTED_MAX_DETAIL_CHECKS || "50", 10);
const WALLET_DETAIL_CHECKS = Number.parseInt(process.env.VINTED_WALLET_DETAIL_CHECKS || "120", 10);
const HISTORY_OLDER_STREAK_STOP = Number.parseInt(process.env.VINTED_HISTORY_OLDER_STREAK_STOP || "5", 10);
const ITEM_QUERY = normalizeTitle(CLI_ITEM_QUERY || process.env.VINTED_ITEM_QUERY || "");

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

async function collectOrders(page, targetCounts, ordersUrl = ORDERS_URL) {
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
        await page.goto(ordersUrl, { waitUntil: "networkidle2" });
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

function parseTargetMonth(value) {
  const match = String(value || "").trim().match(/^(\d{4})-(\d{2})$/);
  if (!match) return "";
  return `${match[1]}-${match[2]}`;
}

function walletHistoryUrlForMonth(targetMonth) {
  const normalized = parseTargetMonth(targetMonth) || currentMonthString();
  const [year, month] = normalized.split("-");
  return `${WALLET_HISTORY_BASE_URL}/${year}/${Number.parseInt(month, 10)}`;
}

function parsePolishWalletDate(value) {
  const text = String(value || "").replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim().toLowerCase();
  if (!text) return "";
  const match = text.match(/^(\d{1,2})\s+([a-ząćęłńóśżź]+)\s+(\d{4})$/i);
  if (!match) return "";
  const monthMap = {
    stycznia: "01",
    lutego: "02",
    marca: "03",
    kwietnia: "04",
    maja: "05",
    czerwca: "06",
    lipca: "07",
    sierpnia: "08",
    wrzesnia: "09",
    września: "09",
    pazdziernika: "10",
    października: "10",
    listopada: "11",
    grudnia: "12",
  };
  const dd = match[1].padStart(2, "0");
  const mm = monthMap[match[2]] || "";
  const yyyy = match[3];
  if (!mm) return "";
  return `${dd}.${mm}.${yyyy}`;
}

function parseDateFromTitleAttr(value) {
  const text = String(value || "");
  const match = text.match(/(\d{1,2})\.(\d{1,2})\.(\d{4})(?:,\s*(\d{1,2}:\d{2})(?::\d{2})?)?/);
  if (!match) return "";
  const dd = match[1].padStart(2, "0");
  const mm = match[2].padStart(2, "0");
  const yyyy = match[3];
  const hhmm = match[4] ? match[4].padStart(5, "0") : "00:00";
  return `${dd}.${mm}.${yyyy} ${hhmm}`;
}

function formatDateForPoland(date) {
  const dd = String(date.getDate()).padStart(2, "0");
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const yyyy = String(date.getFullYear());
  const hh = String(date.getHours()).padStart(2, "0");
  const mi = String(date.getMinutes()).padStart(2, "0");
  return `${dd}.${mm}.${yyyy} ${hh}:${mi}`;
}

function parseRelativePolishDate(relativeText) {
  const text = String(relativeText || "").toLowerCase().replace(/\s+/g, " ").trim();
  if (!text) return "";
  const now = new Date();
  const date = new Date(now.getTime());

  if (text.includes("dzis")) {
    return formatDateForPoland(date);
  }
  if (text.includes("wczoraj")) {
    date.setDate(date.getDate() - 1);
    return formatDateForPoland(date);
  }

  const minuteMatch = text.match(/(\d+)\s*(min|minut)/);
  if (minuteMatch) {
    date.setMinutes(date.getMinutes() - Number.parseInt(minuteMatch[1], 10));
    return formatDateForPoland(date);
  }

  const hourMatch = text.match(/(\d+)\s*(godz|godzin|godziny)/);
  if (hourMatch) {
    date.setHours(date.getHours() - Number.parseInt(hourMatch[1], 10));
    return formatDateForPoland(date);
  }

  const dayMatch = text.match(/(\d+)\s*(dzien|dni|dzień)/);
  if (dayMatch) {
    date.setDate(date.getDate() - Number.parseInt(dayMatch[1], 10));
    return formatDateForPoland(date);
  }

  const weekMatch = text.match(/(\d+)\s*(tydzien|tygodnie|tygodni|tydzień)/);
  if (weekMatch) {
    date.setDate(date.getDate() - Number.parseInt(weekMatch[1], 10) * 7);
    return formatDateForPoland(date);
  }

  const monthMatch = text.match(/(\d+)\s*(miesiac|miesiace|miesiecy|miesiąc|miesiące|miesięcy)/);
  if (monthMatch) {
    date.setMonth(date.getMonth() - Number.parseInt(monthMatch[1], 10));
    return formatDateForPoland(date);
  }

  return "";
}

function parseRelativeDateFromBodyText(text) {
  const value = String(text || "");
  if (!value) return "";
  const pattern =
    /(\d+\s*(?:min|minut|godz|godzin|godziny|dzien|dni|dzień|tydzien|tygodnie|tygodni|tydzień|miesiac|miesiace|miesiecy|miesiąc|miesiące|miesięcy)\s*temu|dzis(?:iaj)?|wczoraj)/gi;
  const matches = value.match(pattern);
  if (!matches || !matches.length) return "";
  // Prefer the last visible relative timestamp from conversation timeline.
  for (let i = matches.length - 1; i >= 0; i -= 1) {
    const parsed = parseRelativePolishDate(matches[i]);
    if (parsed) return parsed;
  }
  return "";
}

function normalizeSimple(value) {
  const replacements = {
    ą: "a",
    ć: "c",
    ę: "e",
    ł: "l",
    ń: "n",
    ó: "o",
    ś: "s",
    ż: "z",
    ź: "z",
  };
  return String(value || "")
    .toLowerCase()
    .replace(/[ąćęłńóśżź]/g, (ch) => replacements[ch] || ch)
    .replace(/\s+/g, " ")
    .trim();
}

function monthFromDateText(value) {
  const match = String(value || "").match(/^(\d{2})\.(\d{2})\.(\d{4})/);
  if (!match) return "";
  return `${match[3]}-${match[2]}`;
}

function isOlderThanTargetMonth(dateText, targetMonth) {
  const itemMonth = monthFromDateText(dateText);
  if (!itemMonth || !targetMonth) return false;
  return itemMonth < targetMonth;
}

async function gotoWithFallback(page, url) {
  try {
    await page.goto(url, { waitUntil: "domcontentloaded", timeout: 45000 });
    return true;
  } catch (err) {
    const msg = String(err && err.message ? err.message : err);
    console.log(`[warn] Problem z otwarciem: ${url} -> ${msg}`);
    return false;
  }
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

async function openCompletedOrders(page, orderType) {
  const url = `${ORDERS_URL}?order_type=${orderType}`;
  await page.goto(url, { waitUntil: "networkidle2" });
  await page.waitForSelector(COMPLETED_FILTER, { timeout: 300000 });
  await page.click(COMPLETED_FILTER);
  await sleep(1500);
  await page.waitForSelector(ORDER_ITEM, { timeout: 30000 });
}

async function extractDetailDate(page) {
  for (let attempt = 0; attempt < 7; attempt += 1) {
    const picked =
      (await page
        .evaluate(() => {
      const hasDate = (value) => /\d{1,2}\.\d{1,2}\.\d{4}/.test(String(value || ""));
      const isStatusNode = (node) => {
        const text = normalizeSimple(node.textContent || "");
        return (
          text.includes("zakup zako") ||
          text.includes("sprzedaz zako") ||
          text.includes("zamowienie zostalo zrealizowane")
        );
      };
      const relativeRegex =
        /(\d+\s*(min|minut|godz|godzin|godziny|dzien|dni|dzień|tydzien|tygodnie|tygodni|tydzień|miesiac|miesiace|miesiecy|miesiąc|miesiące|miesięcy)\s*temu|dzis|wczoraj)/i;

      const pickByNodeDistance = (baseNodes, dateNodes) => {
        let bestTitle = "";
        let bestRelative = "";
        let bestDistance = Number.POSITIVE_INFINITY;
        for (const baseNode of baseNodes) {
          const baseRect = baseNode.getBoundingClientRect();
          const baseTop = baseRect.top;
          for (const dateNode of dateNodes) {
            const top = Number.isFinite(dateNode.top) ? dateNode.top : 0;
            const distance = baseTop - top;
            if (distance >= -30 && distance < bestDistance) {
              bestDistance = distance;
              bestTitle = dateNode.title;
              bestRelative = dateNode.relative;
            }
          }
        }
        return { title: bestTitle, relative: bestRelative };
      };

      const pickNearestPreviousByDomOrder = (statusNode) => {
        const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_ELEMENT);
        let bestTitle = "";
        let bestRelative = "";
        while (walker.nextNode()) {
          const node = walker.currentNode;
          if (node === statusNode) {
            break;
          }
          if (!(node instanceof Element)) continue;
          if (node.matches("span[title]")) {
            const title = node.getAttribute("title") || "";
            if (hasDate(title)) {
              bestTitle = title;
              bestRelative = (node.textContent || "").trim();
            }
          } else if (node.matches("span.web_ui__Text__caption span")) {
            const text = (node.textContent || "").trim();
            if (relativeRegex.test(text)) {
              bestRelative = text;
            }
          }
        }
        return { title: bestTitle, relative: bestRelative };
      };

      const findDateInPreviousSiblings = (startNode) => {
        let cursor = startNode;
        for (let depth = 0; depth < 10 && cursor; depth += 1) {
          let sib = cursor.previousElementSibling;
          while (sib) {
            const dateSpan = sib.querySelector("span[title]");
            if (dateSpan) {
              const title = dateSpan.getAttribute("title") || "";
              if (hasDate(title)) {
                return { title, relative: (dateSpan.textContent || "").trim() };
              }
            }
            const sibText = (sib.textContent || "").trim();
            const relMatch = sibText.match(relativeRegex);
            if (relMatch && relMatch[0]) {
              return { title: "", relative: relMatch[0] };
            }
            sib = sib.previousElementSibling;
          }
          cursor = cursor.parentElement;
        }
        return null;
      };

      const statusNodes = Array.from(
        document.querySelectorAll('[data-testid="conversation-message--status-message"]')
      );
      const successNodes = statusNodes.filter((node) => isStatusNode(node)).reverse();

      const dateNodes = Array.from(document.querySelectorAll("span[title]"))
        .map((node) => ({
          title: node.getAttribute("title") || "",
          relative: (node.textContent || "").trim(),
          top: node.getBoundingClientRect().top,
        }))
        .filter((row) => hasDate(row.title));

      if (successNodes.length) {
        for (const statusNode of successNodes) {
          const bySibling = findDateInPreviousSiblings(statusNode);
          if (bySibling && (bySibling.title || bySibling.relative)) {
            return bySibling;
          }
          const byDomOrder = pickNearestPreviousByDomOrder(statusNode);
          if (byDomOrder.title || byDomOrder.relative) {
            return byDomOrder;
          }
        }
      }

      if (successNodes.length && dateNodes.length) {
        const pickedDate = pickByNodeDistance(successNodes, dateNodes);
        if (pickedDate.title) return pickedDate;
      }

      // Fallback: look for caption "X czasu temu" near status section.
      if (successNodes.length) {
        const relativeNodes = Array.from(document.querySelectorAll("span.web_ui__Text__caption span"))
          .map((node) => ({
            title: node.getAttribute("title") || "",
            relative: (node.textContent || "").trim(),
            top: node.getBoundingClientRect().top,
          }))
          .filter((row) => row.relative);
        if (relativeNodes.length) {
          const pickedRelative = pickByNodeDistance(successNodes, relativeNodes);
          if (pickedRelative.relative) return pickedRelative;
        }
      }

      // Strict mode: do not pick global timestamps from unrelated parts of page.
      return { title: "", relative: "" };
        })
        .catch(() => null)) || { title: "", relative: "" };

    const parsedByTitle = parseDateFromTitleAttr(picked.title || "");
    if (parsedByTitle) return parsedByTitle;
    const parsedByRelative = parseRelativePolishDate(picked.relative || "");
    if (parsedByRelative) return parsedByRelative;

    const bodyText = await page.evaluate(() => document.body.innerText || "");
    const parsedByBody = parseDateFromTitleAttr(bodyText);
    if (parsedByBody) return parsedByBody;

    // Some conversation timelines lazily render date spans until scrolled up.
    await page.evaluate((step) => {
      window.scrollBy(0, -step);
      const scrollables = Array.from(document.querySelectorAll("*")).filter(
        (el) => el.scrollHeight > el.clientHeight + 20
      );
      for (const el of scrollables) {
        el.scrollTop = Math.max(0, el.scrollTop - step);
      }
    }, 700);
    await sleep(350);
  }
  return "";
}

async function collectHistoryForOrderType(page, orderType, targetMonth) {
  await openCompletedOrders(page, orderType);
  const ordersUrl = `${ORDERS_URL}?order_type=${orderType}`;
  const collectedOrders = await collectOrders(page, new Map(), ordersUrl);
  const visibleOrders = ITEM_QUERY
    ? collectedOrders.filter((order) => normalizeTitle(order.title || "").includes(ITEM_QUERY))
    : collectedOrders;
  console.log(`Karty '${orderType}' do sprawdzenia: ${visibleOrders.length}`);
  if (ITEM_QUERY) {
    console.log(`[${orderType}] Filtr item='${ITEM_QUERY}', znaleziono: ${visibleOrders.length}`);
  }

  const results = [];
  let olderStreak = 0;
  for (let idx = 0; idx < visibleOrders.length; idx += 1) {
    const order = visibleOrders[idx];
    if (!order.href) continue;
    const url = order.href.startsWith("http") ? order.href : `https://www.vinted.pl${order.href}`;
    const shortTitle = String(order.title || "").slice(0, 90);
    console.log(
      `[${orderType}] ${idx + 1}/${visibleOrders.length} -> ${shortTitle || "(brak tytulu)"} | ${
        order.price || "(brak ceny)"
      }`
    );

    const opened = await gotoWithFallback(page, url);
    if (!opened) {
      console.log(`[${orderType}] pominieto (blad otwarcia): ${url}`);
      continue;
    }
    await sleep(1200);

    const dateText = await extractDetailDate(page);
    const pageText = await page.evaluate(() => document.body.innerText || "");
    const txNumber = extractTransactionNumber(pageText);
    const locationLabel = await getLocationLabel(page);
    const country = extractCountry(locationLabel);

    if (isOlderThanTargetMonth(dateText, targetMonth)) {
      olderStreak += 1;
      console.log(
        `[${orderType}] starszy miesiac (${dateText || "brak daty"}), streak=${olderStreak}/${HISTORY_OLDER_STREAK_STOP}`
      );
      if (olderStreak >= HISTORY_OLDER_STREAK_STOP) {
        console.log(
          `Wykryto ${olderStreak} starszych pozycji z rzedu dla '${orderType}', koncze ten typ.`
        );
        break;
      }
      continue;
    }
    olderStreak = 0;

    console.log(
      `[${orderType}] OK data=${dateText || "brak"} tx=${txNumber || "-"} kraj=${country || "-"}`
    );
    results.push({
      kind: orderType === "sold" ? "sprzedaz" : "zakup",
      title: order.title || "",
      price: order.price || "",
      date_text: dateText || "",
      country,
      transaction_number: txNumber,
      order_url: url,
      order_type: orderType,
    });
  }

  return results;
}

async function collectWalletHistory(page, targetMonth, orderTypes) {
  const url = walletHistoryUrlForMonth(targetMonth);
  console.log(`Otwieram historie portfela: ${url}`);
  const opened = await gotoWithFallback(page, url);
  if (!opened) {
    throw new Error("Nie udalo sie otworzyc strony historii portfela.");
  }

  const itemSelector = "li.pile__element a.web_ui__Cell__link";
  await page.waitForSelector(itemSelector, { timeout: 30000 });

  const seen = new Map();
  let stagnant = 0;
  for (let i = 0; i < MAX_SCROLLS; i += 1) {
    const items = await page.$$eval(itemSelector, (nodes) =>
      nodes.map((node) => {
        const titleEl = node.querySelector(".web_ui__Cell__title");
        const bodyEl = node.querySelector(".web_ui__Cell__body");
        const amountEl = node.querySelector(".web_ui__Cell__suffix h2");
        const suffixDiv = node.querySelector(".web_ui__Cell__suffix div");
        const href = node.getAttribute("href") || "";
        const amount = (amountEl ? amountEl.textContent : "").replace(/\u00a0/g, " ").trim();
        const suffixText = (suffixDiv ? suffixDiv.textContent : "").replace(/\u00a0/g, " ").trim();
        const dateLabel = suffixText.replace(amount, "").replace(/\s+/g, " ").trim();
        return {
          href,
          kind_label: titleEl ? titleEl.textContent.trim() : "",
          title: bodyEl ? bodyEl.textContent.trim() : "",
          amount,
          date_label: dateLabel,
        };
      })
    );

    const before = seen.size;
    for (const item of items) {
      const key = `${item.href}|${item.kind_label}|${item.title}|${item.amount}|${item.date_label}`;
      if (!seen.has(key)) {
        seen.set(key, item);
      }
    }

    if (seen.size === before) {
      stagnant += 1;
    } else {
      stagnant = 0;
    }

    console.log(`Wallet scroll ${i + 1}/${MAX_SCROLLS}: wpisy=${seen.size}`);
    if (stagnant >= MAX_STAGNANT) {
      console.log("Brak nowych wpisow portfela, koncze przewijanie.");
      break;
    }

    await page.evaluate("window.scrollTo(0, document.body.scrollHeight)");
    await sleep(1200);
  }

  const wantedKinds = new Set();
  const wantedOrderTypes = orderTypes && orderTypes.length ? orderTypes : ["purchased", "sold"];
  for (const type of wantedOrderTypes) {
    if (type === "sold") {
      wantedKinds.add("sprzedaz");
    } else if (type === "purchased") {
      wantedKinds.add("zakup");
      wantedKinds.add("usluga");
      wantedKinds.add("zwrot");
    }
  }

  const entries = Array.from(seen.values());
  const filtered = [];
  for (const entry of entries) {
    const kindText = normalizeSimple(entry.kind_label || "");
    let kind = "";
    if (kindText.includes("sprzedane")) {
      kind = "sprzedaz";
    } else if (kindText.includes("zwrot")) {
      kind = "zwrot";
    } else if (kindText.includes("zakup") && kindText.includes("ekspozycja")) {
      kind = "usluga";
    } else if (kindText.includes("zakup")) {
      kind = "zakup";
    } else {
      continue;
    }
    if (!wantedKinds.has(kind)) {
      continue;
    }
    if (ITEM_QUERY && !normalizeTitle(entry.title || "").includes(ITEM_QUERY)) {
      continue;
    }
    const dateText = parsePolishWalletDate(entry.date_label);
    if (targetMonth && dateText) {
      const monthKey = monthFromDateText(dateText);
      if (monthKey && monthKey !== targetMonth) {
        continue;
      }
    }
    filtered.push({
      kind,
      title: entry.title || "",
      price: entry.amount || "",
      date_text: dateText,
      country: "",
      transaction_number: "",
      order_url: entry.href ? (entry.href.startsWith("http") ? entry.href : `https://www.vinted.pl${entry.href}`) : "",
      order_type: kind === "sprzedaz" ? "sold" : kind === "usluga" ? "service" : "purchased",
    });
  }

  console.log(`Po filtrach wallet: ${filtered.length}`);
  for (let i = 0; i < filtered.length; i += 1) {
    const row = filtered[i];
    console.log(
      `[wallet] ${i + 1}/${filtered.length} ${row.kind} | ${row.title || "(brak tytulu)"} | ${
        row.price || "(brak kwoty)"
      } | ${row.date_text || "brak daty"}`
    );
  }

  const salesWithLinks = filtered.filter((row) => row.kind === "sprzedaz" && row.order_url);
  if (!salesWithLinks.length) {
    return filtered;
  }

  const maxChecks = Math.min(salesWithLinks.length, Math.max(0, WALLET_DETAIL_CHECKS));
  console.log(`Uzupelniam numery transakcji z detali wallet: ${maxChecks}/${salesWithLinks.length}`);
  for (let i = 0; i < maxChecks; i += 1) {
    const row = salesWithLinks[i];
    const opened = await gotoWithFallback(page, row.order_url);
    if (!opened) {
      console.log(`[wallet-details] pominieto (blad otwarcia): ${row.order_url}`);
      continue;
    }
    await sleep(900);
    const pageText = await page.evaluate(() => document.body.innerText || "");
    const txNumber = extractTransactionNumber(pageText);
    if (txNumber) {
      row.transaction_number = txNumber;
    }
    if (!row.country) {
      const locationLabel = await getLocationLabel(page);
      row.country = extractCountry(locationLabel);
    }
    console.log(
      `[wallet-details] ${i + 1}/${maxChecks} tx=${row.transaction_number || "-"} kraj=${row.country || "-"}`
    );
  }

  return filtered;
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
  if (PROFILE_NAME) {
    launchOptions.args.push(`--profile-directory=${PROFILE_NAME}`);
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

  if (SCRAPE_MODE === "history") {
    const targetMonth = parseTargetMonth(HISTORY_TARGET_MONTH) || currentMonthString();
    const orderTypes = HISTORY_ORDER_TYPES.length ? HISTORY_ORDER_TYPES : ["purchased", "sold"];
    console.log(`Tryb historyczny (wallet). Miesiac docelowy: ${targetMonth}. Typy: ${orderTypes.join(", ")}`);
    const all = await collectWalletHistory(page, targetMonth, orderTypes);
    console.log(JSON.stringify(all));
    await browser.close();
    return;
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
