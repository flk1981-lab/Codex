/*
  ChatGPT history exporter for your own account.

  How to use:
  1. Open https://chatgpt.com/ while logged in.
  2. Press F12 to open DevTools.
  3. Open the Console tab.
  4. Paste this entire script and press Enter.
  5. Wait for the download prompt.

  Output:
  - One JSON file containing:
    - export metadata
    - conversation summaries
    - full conversation payloads returned by ChatGPT

  Notes:
  - This script relies on your current browser session and only exports data
    that your logged-in account can access.
  - Internal endpoints may change in the future. If that happens, the script
    will need a small update.
*/

(async () => {
  const CONFIG = {
    pageSize: 100,
    pauseMs: 350,
    maxRetries: 3,
    includeArchived: true,
    filenamePrefix: "chatgpt-history-export",
  };

  const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

  async function fetchJson(url, attempt = 1) {
    const response = await fetch(url, {
      method: "GET",
      credentials: "include",
      headers: {
        "accept": "application/json",
      },
    });

    if (!response.ok) {
      const text = await response.text().catch(() => "");
      const message = `HTTP ${response.status} for ${url}\n${text.slice(0, 500)}`;

      if (attempt < CONFIG.maxRetries && response.status >= 500) {
        console.warn(`Retry ${attempt} failed, trying again: ${url}`);
        await sleep(CONFIG.pauseMs * attempt);
        return fetchJson(url, attempt + 1);
      }

      throw new Error(message);
    }

    return response.json();
  }

  async function fetchConversationPage(offset) {
    const url =
      `/backend-api/conversations?offset=${offset}` +
      `&limit=${CONFIG.pageSize}&order=updated`;
    return fetchJson(url);
  }

  async function fetchConversationDetail(conversationId) {
    const url = `/backend-api/conversation/${encodeURIComponent(conversationId)}`;
    return fetchJson(url);
  }

  function downloadJson(filename, data) {
    const blob = new Blob([JSON.stringify(data, null, 2)], {
      type: "application/json",
    });
    const href = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = href;
    anchor.download = filename;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    setTimeout(() => URL.revokeObjectURL(href), 1000);
  }

  const startedAt = new Date().toISOString();
  console.log(`[export] Started at ${startedAt}`);

  const summaries = [];
  const fullConversations = [];
  const failedDetails = [];

  let offset = 0;
  let total = null;

  while (true) {
    const page = await fetchConversationPage(offset);
    const items = Array.isArray(page.items) ? page.items : [];

    if (typeof page.total === "number") {
      total = page.total;
    }

    const filteredItems = CONFIG.includeArchived
      ? items
      : items.filter((item) => !item.is_archived);

    summaries.push(...filteredItems);

    console.log(
      `[export] Loaded summaries ${summaries.length}` +
        (total !== null ? ` / ${total}` : "")
    );

    if (items.length < CONFIG.pageSize) {
      break;
    }

    offset += items.length;
    await sleep(CONFIG.pauseMs);
  }

  for (let index = 0; index < summaries.length; index += 1) {
    const summary = summaries[index];
    const label = `${index + 1}/${summaries.length}`;

    try {
      console.log(`[export] Fetching detail ${label}: ${summary.title || summary.id}`);
      const detail = await fetchConversationDetail(summary.id);
      fullConversations.push({
        summary,
        detail,
      });
    } catch (error) {
      console.error(`[export] Failed detail ${label}: ${summary.id}`, error);
      failedDetails.push({
        id: summary.id,
        title: summary.title || null,
        error: String(error),
      });
    }

    await sleep(CONFIG.pauseMs);
  }

  const finishedAt = new Date().toISOString();
  const stamp = startedAt.replace(/[:.]/g, "-");
  const filename = `${CONFIG.filenamePrefix}-${stamp}.json`;

  const payload = {
    exported_at: finishedAt,
    started_at: startedAt,
    source_origin: location.origin,
    summary_count: summaries.length,
    detail_count: fullConversations.length,
    failed_detail_count: failedDetails.length,
    conversations: fullConversations,
    failed_details: failedDetails,
  };

  downloadJson(filename, payload);

  console.log(`[export] Finished at ${finishedAt}`);
  console.log(`[export] Downloaded file: ${filename}`);
  console.log(
    `[export] Success: ${fullConversations.length}, Failed: ${failedDetails.length}`
  );
})();
