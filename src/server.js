const path = require('path');
const fs = require('fs');
const express = require('express');

const config = require('./config');
const { getProjects, getProjectUsers, getProjectClientUsers } = require('./zohoClient');
const { sanitizeMomPayload, validateMomPayload } = require('./momTemplate');
const { generateMomPdf } = require('./pdfService');

const app = express();

if (!fs.existsSync(config.app.generatedDir)) {
  fs.mkdirSync(config.app.generatedDir, { recursive: true });
}

app.use(express.json({ limit: '5mb' }));
app.use(express.urlencoded({ extended: true }));

app.use('/generated-pdfs', express.static(config.app.generatedDir));
app.use(express.static(path.join(config.app.rootDir, 'public')));

function buildAbsolutePdfUrl(pdfUrl) {
  try {
    return new URL(pdfUrl, config.app.baseUrl).toString();
  } catch (_error) {
    return `${String(config.app.baseUrl || '').replace(/\/+$/, '')}${pdfUrl}`;
  }
}

function buildOutlookDraft({ mom, options, pdfUrl }) {
  const to = String(options.emailTo || '').trim();
  const cc = String(options.emailCc || '').trim();
  const subject = String(options.emailSubject || '').trim() || `Minutes of Meeting - ${mom.projectName}`;
  const pdfAbsoluteUrl = buildAbsolutePdfUrl(pdfUrl);

  const defaultBodyLines = [
    'Dear Team,',
    '',
    `Please find the Minutes of Meeting details for "${mom.projectName}" below:`,
    `- Meeting Title: ${mom.meetingTitle || '-'}`,
    `- Project: ${mom.projectName || '-'}`,
    `- Date: ${mom.meetingDate || '-'}`,
    `- Time: ${mom.meetingTime || '-'}`,
    `- Location: ${mom.meetingLocation || '-'}`,
    '',
    `PDF Link: ${pdfAbsoluteUrl}`,
    '',
    'Note: Due to browser and Outlook security limits, the PDF cannot be auto-attached by web link.',
    'Please attach the downloaded PDF manually before sending.',
    '',
    'Regards,',
    'M.O.M App'
  ];

  const customBody = String(options.emailBody || '').trim();
  const body = customBody
    ? `${customBody}\n\nPDF Link: ${pdfAbsoluteUrl}\n\nPlease attach the generated PDF manually before sending.`
    : defaultBodyLines.join('\n');

  const outlookParams = new URLSearchParams();
  if (to) {
    outlookParams.set('to', to);
  }
  if (cc) {
    outlookParams.set('cc', cc);
  }
  outlookParams.set('subject', subject);
  outlookParams.set('body', body);

  const mailtoParams = new URLSearchParams();
  if (cc) {
    mailtoParams.set('cc', cc);
  }
  mailtoParams.set('subject', subject);
  mailtoParams.set('body', body);

  return {
    mode: 'outlook-draft',
    to,
    cc,
    subject,
    body,
    pdfAbsoluteUrl,
    outlookComposeUrl: `https://outlook.office.com/mail/deeplink/compose?${outlookParams.toString()}`,
    mailtoUrl: `mailto:${encodeURIComponent(to)}?${mailtoParams.toString()}`,
    attachmentAutoSupported: false,
    attachmentNote:
      'Attachment cannot be auto-added by browser deeplink. The generated PDF is opened for manual attachment.'
  };
}

app.get('/api/health', (_req, res) => {
  const zohoAutoRefreshConfigured = Boolean(
    config.zoho.refreshToken && config.zoho.clientId && config.zoho.clientSecret
  );

  res.json({
    success: true,
    timestamp: new Date().toISOString(),
    zohoMode: config.zoho.useMock ? 'mock' : 'live',
    zohoAutoRefreshConfigured,
    emailEnabled: true,
    emailMode: 'outlook-draft'
  });
});

app.get('/api/zoho/projects', async (req, res) => {
  try {
    const query = String(req.query.query || '').trim();
    const projects = await getProjects(query);

    res.json({
      success: true,
      projects
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      message: error.message || 'Failed to fetch projects from Zoho Projects'
    });
  }
});

app.get('/api/zoho/projects/:projectId/users', async (req, res) => {
  try {
    const projectId = String(req.params.projectId || '').trim();
    const users = await getProjectUsers(projectId);

    res.json({
      success: true,
      users
    });
  } catch (error) {
    const projectId = String(req.params.projectId || '').trim();
    res.status(500).json({
      success: false,
      message:
        error.message || `Failed to fetch Zoho project users for project ID: ${projectId || '(empty)'}`
    });
  }
});

app.get('/api/zoho/projects/:projectId/client-users', async (req, res) => {
  try {
    const projectId = String(req.params.projectId || '').trim();
    const users = await getProjectClientUsers(projectId);

    res.json({
      success: true,
      users
    });
  } catch (error) {
    res.json({
      success: true,
      users: [],
      message: error.message || 'Zoho client users not available for this project.'
    });
  }
});

app.post('/api/mom/submit', async (req, res) => {
  try {
    const mom = sanitizeMomPayload(req.body.mom || {});
    const options = req.body.options || {};

    const validationErrors = validateMomPayload(mom);
    if (validationErrors.length) {
      return res.status(400).json({
        success: false,
        message: 'Validation failed',
        errors: validationErrors
      });
    }

    const shouldGeneratePdf = Boolean(options.generatePdf || options.printPdf || options.sendEmail);

    if (!shouldGeneratePdf) {
      return res.status(400).json({
        success: false,
        message: 'Select at least one output option (Email, PDF, Print).'
      });
    }

    const { fileName } = await generateMomPdf(mom, config.app.generatedDir);
    const pdfUrl = `/generated-pdfs/${fileName}`;

    let emailDraft = null;
    if (options.sendEmail) {
      emailDraft = buildOutlookDraft({
        mom,
        options,
        pdfUrl
      });
    }

    return res.json({
      success: true,
      message: 'M.O.M processed successfully.',
      result: {
        pdfUrl,
        pdfFileName: fileName,
        pdfAbsoluteUrl: buildAbsolutePdfUrl(pdfUrl),
        emailSent: false,
        emailDraft,
        printRequested: Boolean(options.printPdf)
      }
    });
  } catch (error) {
    return res.status(500).json({
      success: false,
      message: error.message || 'Failed to process M.O.M submission'
    });
  }
});

app.get('*', (_req, res) => {
  res.sendFile(path.join(config.app.rootDir, 'public', 'index.html'));
});

const server = app.listen(config.app.port, config.app.host, () => {
  // eslint-disable-next-line no-console
  console.log(`M.O.M app running at ${config.app.baseUrl} on port ${config.app.port}`);
});

server.on('error', (error) => {
  // eslint-disable-next-line no-console
  console.error(`Server startup failed: ${error.code || error.message}`);
  process.exit(1);
});
