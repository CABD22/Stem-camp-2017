# Cloudflare Worker Review

This document captures feedback on the Financial Planning Email → Asana Task worker implementation that was shared for review.

## Summary

The worker already covers the major flows—handling Asana webhook handshakes, accepting POST payloads from Outlook, and creating follow-up tasks in Asana. A few refinements will make it more robust and align the Asana API calls with their expected formats.

## Recommendations

1. **Only send the date custom field when a real date is available**  
   When the Outlook payload omits `date_sent`, the worker currently sends `{ "date": "Not provided" }`. Asana treats that as an invalid value and rejects the request. Build the `custom_fields` object conditionally so that the date field is set only when you have a properly formatted ISO date.

2. **Validate the inbound payload up front**  
   Guard clauses that check for a non-empty recipient and subject will let you return a `400 Bad Request` early instead of reaching Asana with placeholders such as “Unknown.” This also keeps accidental or malicious requests from creating placeholder tasks.

3. **Improve error logging for Asana responses**  
   On failures, capturing `await asanaResp.text()` (rather than just `json()`) in the error message ensures you see the raw payload even when the response body is not JSON (for example, on HTML error pages). Logging the response status and body to Workers KV or console output can speed up debugging.

4. **Protect the endpoint with a shared secret**  
   Because the worker accepts arbitrary POSTs, add a simple authentication check—e.g., require an `X-Worker-Token` header that matches an environment secret—so only your Outlook automation can create tasks.

5. **Normalize `name` field usage**  
   The current payload maps `data.name` into the task notes, but the field is optional for the automation. If `name` is truly required, validate it the same way you validate email and subject; otherwise, consider removing it from the notes to avoid storing the literal word “Unknown.”

6. **Handle webhook endpoint separately (optional)**  
   If you don’t plan to receive webhook events yet, you can remove the `/asana-webhook` branch until it is needed. This reduces the surface area and keeps the worker focused on the send-to-task flow. When you do add webhook handling, ensure the POST body is validated and that you respond quickly to avoid retries.

## Suggested Code Adjustments

Below is a sketch of how you might incorporate the most critical changes (payload validation and conditional date handling):

```javascript
const requiredFields = ["email", "subject"];
for (const field of requiredFields) {
  if (!data[field]) {
    return new Response(`Missing required field: ${field}`, { status: 400 });
  }
}

const customFields = {
  "1209562543626183": data.email,
  "1211619397148241": data.subject
};
if (data.date_sent && /^\d{4}-\d{2}-\d{2}$/.test(data.date_sent)) {
  customFields["1211619397148231"] = { date: data.date_sent };
}
```

With these adjustments, the worker will only create tasks when the essential information is present, and it will avoid sending invalid date values to Asana.
