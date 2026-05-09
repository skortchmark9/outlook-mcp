/**
 * Reply to email functionality (preserves threading)
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

function parseRecipients(value) {
  if (!value) return undefined;
  return value.split(',').map(email => ({
    emailAddress: { address: email.trim() }
  }));
}

async function handleReplyEmail(args) {
  const { id, body, replyAll = false, to, cc, isHtml } = args;

  if (!id) {
    return { content: [{ type: "text", text: "Email id is required." }] };
  }
  if (!body) {
    return { content: [{ type: "text", text: "Body content is required." }] };
  }

  try {
    const accessToken = await ensureAuthenticated();

    const contentType = isHtml === true ? 'html'
      : isHtml === false ? 'text'
      : (body.includes('<html') || body.includes('<HTML')) ? 'html' : 'text';

    const message = {
      body: { contentType, content: body }
    };
    const toRecipients = parseRecipients(to);
    const ccRecipients = parseRecipients(cc);
    if (toRecipients) message.toRecipients = toRecipients;
    if (ccRecipients) message.ccRecipients = ccRecipients;

    const action = replyAll ? 'replyAll' : 'reply';
    const endpoint = `me/messages/${encodeURIComponent(id)}/${action}`;

    await callGraphAPI(accessToken, 'POST', endpoint, { message });

    return {
      content: [{
        type: "text",
        text: `Reply sent successfully (${action}).`
      }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{ type: "text", text: "Authentication required. Please use the 'authenticate' tool first." }]
      };
    }
    return {
      content: [{ type: "text", text: `Error sending reply: ${error.message}` }]
    };
  }
}

module.exports = handleReplyEmail;
