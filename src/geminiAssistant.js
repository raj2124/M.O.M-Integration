const { GoogleGenerativeAI } = require('@google/generative-ai');

const SUPPORTED_MIME_TYPES = new Set([
  'application/pdf',
  'image/png',
  'image/jpeg',
  'image/jpg',
  'image/webp'
]);

function ensureGeminiConfigured(config) {
  if (!config?.gemini?.enabled) {
    throw new Error('Gemini assistant is disabled.');
  }

  if (!String(config.gemini.apiKey || '').trim()) {
    throw new Error('Gemini API key is not configured.');
  }
}

function assertSupportedFile(file) {
  if (!file) {
    throw new Error('Please upload an image or PDF file.');
  }

  const mimeType = String(file.mimetype || '').trim().toLowerCase();
  if (!SUPPORTED_MIME_TYPES.has(mimeType)) {
    throw new Error('Unsupported file type. Please upload PDF, PNG, JPG, JPEG, or WEBP.');
  }

  if (!file.buffer || !Buffer.isBuffer(file.buffer) || file.buffer.length === 0) {
    throw new Error('Uploaded file is empty.');
  }
}

function extractJsonPayload(rawText) {
  const text = String(rawText || '').trim();
  if (!text) {
    throw new Error('Gemini returned an empty response.');
  }

  const fenced = text.match(/```(?:json)?\s*([\s\S]*?)```/i);
  const candidate = fenced ? fenced[1].trim() : text;

  try {
    return JSON.parse(candidate);
  } catch (_error) {
    const firstBrace = candidate.indexOf('{');
    const lastBrace = candidate.lastIndexOf('}');
    if (firstBrace !== -1 && lastBrace !== -1 && lastBrace > firstBrace) {
      return JSON.parse(candidate.slice(firstBrace, lastBrace + 1));
    }
    throw new Error('Gemini response was not valid JSON.');
  }
}

function normalizeBulletBlock(value) {
  const raw = String(value || '').trim();
  if (!raw) {
    return '';
  }

  const lines = raw
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => {
      const cleaned = line.replace(/^[-*•\d.)\s]+/, '').trim();
      return cleaned ? `• ${cleaned}` : '';
    })
    .filter(Boolean);

  return lines.join('\n');
}

function normalizeAssistantPayload(payload = {}) {
  return {
    agenda: normalizeBulletBlock(payload.agenda),
    discussion: normalizeBulletBlock(payload.discussion),
    action_items: normalizeBulletBlock(payload.action_items)
  };
}

function buildPrompt() {
  return [
    'You are an AI assistant for a Minutes of Meeting (M.O.M) SaaS application.',
    'This is a meeting document. Extract ONLY the relevant meeting content.',
    'Ignore headers, footers, logos, stamps, signatures, page numbers, and decorative text.',
    'Return valid JSON only, with this exact shape:',
    '{',
    '  "agenda": "clean bullet points",',
    '  "discussion": "clean bullet points",',
    '  "action_items": "clean bullet points"',
    '}',
    'Rules:',
    '- agenda: extract agenda points only.',
    '- discussion: extract discussion points or topics discussed.',
    '- action_items: extract action items, follow-ups, responsibilities, or next steps.',
    '- If a section is missing, return an empty string for that field.',
    '- Keep text concise and clean.',
    '- Use plain text bullet points separated by new lines.',
    '- Do not return markdown code fences.'
  ].join('\n');
}

async function scanMeetingDocumentWithGemini(config, file) {
  ensureGeminiConfigured(config);
  assertSupportedFile(file);

  const client = new GoogleGenerativeAI(String(config.gemini.apiKey || '').trim());
  const model = client.getGenerativeModel({
    model: String(config.gemini.model || 'gemini-2.5-flash').trim()
  });

  const prompt = buildPrompt();
  const inlineData = {
    mimeType: String(file.mimetype || '').trim(),
    data: file.buffer.toString('base64')
  };

  const result = await model.generateContent({
    contents: [
      {
        role: 'user',
        parts: [{ text: prompt }, { inlineData }]
      }
    ],
    generationConfig: {
      temperature: 0.1,
      responseMimeType: 'application/json'
    }
  });

  const response = await result.response;
  const text = typeof response.text === 'function' ? response.text() : '';
  const parsed = extractJsonPayload(text);
  return normalizeAssistantPayload(parsed);
}

module.exports = {
  scanMeetingDocumentWithGemini
};
