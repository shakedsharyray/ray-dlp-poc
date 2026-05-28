// Ray DLP classifier — Cloudflare Worker
//
// POST /classify
// Body: {
//   sender:      "user@domain",
//   subject:     string,
//   body:        string (text-coerced from message),
//   to:          [{emailAddress, displayName}],
//   cc:          [{emailAddress, displayName}],
//   attachments: [{name, contentType, size}],
//   itemId:      string,
//   timestamp:   ISO8601 string
// }
//
// Response: { action: "block" | "allow", reason: string }

const BLOCK_TERM = "bamba";

export default {
  async fetch(request) {
    if (request.method === "OPTIONS") {
      return new Response(null, { status: 204, headers: corsHeaders() });
    }
    if (request.method !== "POST") {
      return jsonResponse({ error: "POST only" }, 405);
    }

    let msg;
    try {
      msg = await request.json();
    } catch {
      return jsonResponse({ action: "block", reason: "Invalid request payload." }, 400);
    }

    // Log what we received so it shows up in `wrangler tail`.
    // Useful for the user to see exactly what Office.js is passing us.
    console.log("classify request", JSON.stringify({
      sender: msg.sender,
      subject: (msg.subject || "").slice(0, 200),
      bodyChars: (msg.body || "").length,
      toCount: (msg.to || []).length,
      ccCount: (msg.cc || []).length,
      attachmentCount: (msg.attachments || []).length,
      attachments: (msg.attachments || []).map(a => ({
        name: a.name,
        size: a.size,
        contentType: a.contentType,
      })),
      itemId: msg.itemId,
      timestamp: msg.timestamp,
    }));

    const haystack = ((msg.subject || "") + " " + (msg.body || "")).toLowerCase();

    if (haystack.includes(BLOCK_TERM)) {
      return jsonResponse({
        action: "block",
        reason: `Blocked by Ray DLP: this message contains the term '${BLOCK_TERM}'. Please remove it and try again.`,
      });
    }

    return jsonResponse({ action: "allow", reason: "No policy violations detected." });
  },
};

function corsHeaders() {
  return {
    // PoC: permissive. Tighten to specific Outlook origins when productionizing.
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, Authorization",
    "Access-Control-Max-Age": "86400",
  };
}

function jsonResponse(body, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "Content-Type": "application/json", ...corsHeaders() },
  });
}
