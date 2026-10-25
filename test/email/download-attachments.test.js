const fs = require('fs');
const os = require('os');
const path = require('path');
const handleDownloadAttachments = require('../../email/download-attachments');
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');

const FILE = '#microsoft.graph.fileAttachment';

describe('handleDownloadAttachments', () => {
  let outputDir;

  beforeEach(() => {
    callGraphAPI.mockReset();
    ensureAuthenticated.mockResolvedValue('dummy_access_token');
    outputDir = fs.mkdtempSync(path.join(os.tmpdir(), 'attachments-'));
  });

  afterEach(() => {
    fs.rmSync(outputDir, { recursive: true, force: true });
  });

  test('requires id and outputDir', async () => {
    const result = await handleDownloadAttachments({ id: 'msg1' });
    expect(result.content[0].text).toContain('required');
    expect(callGraphAPI).not.toHaveBeenCalled();
  });

  test('saves file attachments, skips inline and item attachments, dedupes names', async () => {
    callGraphAPI.mockImplementation(async (token, method, endpoint) => {
      if (endpoint === 'me/messages/msg1/attachments') {
        return {
          value: [
            { '@odata.type': FILE, id: 'a1', name: 'report.pdf', contentType: 'application/pdf', isInline: false },
            { '@odata.type': FILE, id: 'a2', name: 'report.pdf', contentType: 'application/pdf', isInline: false },
            { '@odata.type': FILE, id: 'a3', name: 'logo.png', contentType: 'image/png', isInline: true },
            { '@odata.type': '#microsoft.graph.itemAttachment', id: 'a4', name: 'Fwd', isInline: false }
          ]
        };
      }
      const id = endpoint.split('/').pop();
      return { contentBytes: Buffer.from(`bytes-${id}`).toString('base64') };
    });

    const result = await handleDownloadAttachments({ id: 'msg1', outputDir });
    const text = result.content[0].text;

    expect(fs.readFileSync(path.join(outputDir, 'report.pdf'), 'utf8')).toBe('bytes-a1');
    expect(fs.readFileSync(path.join(outputDir, 'report (2).pdf'), 'utf8')).toBe('bytes-a2');
    expect(fs.existsSync(path.join(outputDir, 'logo.png'))).toBe(false);
    expect(text).toContain('Saved 2 attachment(s)');
    expect(text).toContain('logo.png (inline)');
    expect(text).toContain('Fwd (#microsoft.graph.itemAttachment)');
  });

  test('strips path components from attachment names', async () => {
    callGraphAPI.mockImplementation(async (token, method, endpoint) => {
      if (endpoint.endsWith('/attachments')) {
        return { value: [{ '@odata.type': FILE, id: 'a1', name: '../../evil.txt', contentType: 'text/plain', isInline: false }] };
      }
      return { contentBytes: Buffer.from('x').toString('base64') };
    });

    await handleDownloadAttachments({ id: 'msg1', outputDir });
    expect(fs.readdirSync(outputDir)).toEqual(['evil.txt']);
  });
});
