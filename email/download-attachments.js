/**
 * Download email attachments to local disk
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

const FILE_ATTACHMENT = '#microsoft.graph.fileAttachment';

/**
 * Turn an attachment name into a safe filename, unique within outputDir
 * @param {string} outputDir - Directory the file will be written to
 * @param {string} name - Attachment name as sent
 * @param {Set<string>} taken - Filenames already used in this call
 * @returns {string} - Absolute path to write to
 */
function uniquePath(outputDir, name, taken) {
  const safe = path.basename(name).replace(/[\/\\:\0]/g, '_') || 'attachment';
  const { name: stem, ext } = path.parse(safe);
  let candidate = safe;
  for (let n = 2; taken.has(candidate) || fs.existsSync(path.join(outputDir, candidate)); n++) {
    candidate = `${stem} (${n})${ext}`;
  }
  taken.add(candidate);
  return path.join(outputDir, candidate);
}

/**
 * Download attachments handler
 * @param {object} args - Tool arguments
 * @param {string} args.id - Email ID (required)
 * @param {string} args.outputDir - Directory to save into (required; created if missing)
 * @param {boolean} args.includeInline - Also save inline attachments (signature images etc.)
 * @returns {object} - MCP response
 */
async function handleDownloadAttachments(args = {}) {
  const emailId = args.id;
  const includeInline = args.includeInline === true;

  if (!emailId || !args.outputDir) {
    return {
      content: [{ type: "text", text: "Both 'id' and 'outputDir' are required." }]
    };
  }

  const outputDir = path.resolve(args.outputDir.replace(/^~(?=$|\/)/, os.homedir()));

  try {
    const accessToken = await ensureAuthenticated();
    const base = `me/messages/${encodeURIComponent(emailId)}/attachments`;

    const listing = await callGraphAPI(accessToken, 'GET', base, null, {
      $select: 'id,name,contentType,size,isInline'
    });
    const attachments = listing.value || [];

    if (attachments.length === 0) {
      return {
        content: [{ type: "text", text: "This email has no attachments." }]
      };
    }

    fs.mkdirSync(outputDir, { recursive: true });

    const saved = [];
    const skipped = [];
    const taken = new Set();

    for (const meta of attachments) {
      if (meta.isInline && !includeInline) {
        skipped.push(`${meta.name} (inline)`);
        continue;
      }
      if (meta['@odata.type'] !== FILE_ATTACHMENT) {
        // Item attachments (attached emails/events) and reference attachments (cloud links) have no bytes
        skipped.push(`${meta.name} (${meta['@odata.type']})`);
        continue;
      }

      const full = await callGraphAPI(accessToken, 'GET', `${base}/${encodeURIComponent(meta.id)}`);
      const buffer = Buffer.from(full.contentBytes, 'base64');
      const target = uniquePath(outputDir, meta.name, taken);
      fs.writeFileSync(target, buffer);
      saved.push(`${target} (${meta.contentType}, ${buffer.length} bytes)`);
    }

    let text = saved.length > 0
      ? `Saved ${saved.length} attachment(s):\n${saved.map(s => `- ${s}`).join('\n')}`
      : 'No attachments saved.';
    if (skipped.length > 0) {
      text += `\n\nSkipped ${skipped.length}:\n${skipped.map(s => `- ${s}`).join('\n')}`;
    }

    return {
      content: [{ type: "text", text }]
    };
  } catch (error) {
    if (error.message === 'Authentication required' || error.message === 'UNAUTHORIZED') {
      return {
        content: [{ type: "text", text: "Authentication required. Please use the 'authenticate' tool first." }]
      };
    }
    return {
      content: [{ type: "text", text: `Failed to download attachments: ${error.message}` }]
    };
  }
}

module.exports = handleDownloadAttachments;
