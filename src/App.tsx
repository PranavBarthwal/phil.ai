import { useEffect, useMemo, useState } from 'react'

type AdSignals = {
  headline: string
  offer: string
  audience: string
  tone: string
  cta: string
}

type ChangeItem = {
  element: string
  original: string
  rewritten: string
}

type PersonalizationOutput = {
  html: string
  personalization_rationale: string
  changes: ChangeItem[]
}

type StepLog = {
  id: number
  time: string
  message: string
  level: 'info' | 'success' | 'error'
}

type ProcessingStage = 'idle' | 'reading' | 'analyzing' | 'personalizing' | 'done'

const apiKeyStorageKey = 'messagematch.geminiApiKey'
const geminiModelCandidates = [
  'gemini-2.5-flash',
  'gemini-2.0-flash',
  'gemini-2.0-flash-lite',
  'gemini-1.5-flash-latest',
  'gemini-1.5-flash',
  'gemini-1.5-pro-latest',
  'gemini-1.5-pro',
]

async function fetchAvailableModels(apiKey: string, onLog?: (message: string) => void) {
  const endpoints = [
    `https://generativelanguage.googleapis.com/v1beta/models?key=${encodeURIComponent(apiKey)}`,
    `https://generativelanguage.googleapis.com/v1/models?key=${encodeURIComponent(apiKey)}`,
  ]

  for (const endpoint of endpoints) {
    try {
      onLog?.(`Checking available models via ${endpoint.includes('/v1beta/') ? 'v1beta/models' : 'v1/models'}...`)
      const response = await fetch(endpoint)
      if (!response.ok) {
        onLog?.(`Model list request failed with HTTP ${response.status}.`)
        continue
      }

      const payload = (await response.json()) as {
        models?: Array<{
          name?: string
          supportedGenerationMethods?: string[]
        }>
      }

      const names = (payload.models ?? [])
        .filter((model) => (model.supportedGenerationMethods ?? []).includes('generateContent'))
        .map((model) => model.name?.replace(/^models\//, '').trim() ?? '')
        .filter(Boolean)

      if (names.length > 0) {
        onLog?.(`Discovered ${names.length} available model(s).`)
        return names
      }
    } catch {
      onLog?.('Model discovery request failed due to network/runtime error.')
      continue
    }
  }

  onLog?.('No discoverable models were returned by the API.')
  return [] as string[]
}

function normalizeUrl(raw: string) {
  const trimmed = raw.trim()
  if (!trimmed) {
    return ''
  }

  return /^https?:\/\//i.test(trimmed) ? trimmed : `https://${trimmed}`
}

function parseJsonResponse<T>(raw: string): T {
  const fenced = raw.match(/```(?:json)?\s*([\s\S]*?)```/i)
  const candidate = (fenced?.[1] ?? raw).trim()

  try {
    return JSON.parse(candidate) as T
  } catch {
    const start = candidate.indexOf('{')
    const end = candidate.lastIndexOf('}')

    if (start >= 0 && end > start) {
      return JSON.parse(candidate.slice(start, end + 1)) as T
    }

    throw new Error('Gemini returned invalid JSON.')
  }
}

async function fileToBase64(file: File) {
  const buffer = await file.arrayBuffer()
  const bytes = new Uint8Array(buffer)
  let binary = ''

  for (const byte of bytes) {
    binary += String.fromCharCode(byte)
  }

  return btoa(binary)
}

async function callGemini<T>(params: {
  apiKey: string
  prompt: string
  imageFile?: File
  onLog?: (message: string) => void
}): Promise<T> {
  const parts: Array<{ text?: string; inlineData?: { mimeType: string; data: string } }> = [
    {
      text: params.prompt,
    },
  ]

  if (params.imageFile) {
    parts.push({
      inlineData: {
        mimeType: params.imageFile.type || 'image/png',
        data: await fileToBase64(params.imageFile),
      },
    })
  }

  const requestBody = {
    contents: [{ role: 'user', parts }],
    generationConfig: {
      temperature: 0.4,
      responseMimeType: 'application/json',
    },
  }

  const failures: string[] = []
  const discoveredModels = await fetchAvailableModels(params.apiKey, params.onLog)
  const modelCandidates = Array.from(new Set([...geminiModelCandidates, ...discoveredModels]))
  params.onLog?.(`Trying ${modelCandidates.length} candidate model endpoint(s).`)

  for (const model of modelCandidates) {
    params.onLog?.(`Calling model: ${model}`)
    const response = await fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${encodeURIComponent(params.apiKey)}`,
      {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(requestBody),
      },
    )

    if (!response.ok) {
      const reason = await response.text()
      failures.push(`${model} -> ${response.status}`)
      params.onLog?.(`Model ${model} failed with HTTP ${response.status}.`)

      if (response.status === 404) {
        continue
      }

      throw new Error(`Gemini request failed (${response.status}): ${reason}`)
    }

    const payload = (await response.json()) as {
      candidates?: Array<{
        content?: { parts?: Array<{ text?: string }> }
      }>
    }

    const text = payload.candidates?.[0]?.content?.parts?.map((part) => part.text ?? '').join('') ?? ''

    if (!text) {
      throw new Error('Gemini returned an empty response.')
    }

    params.onLog?.(`Model ${model} succeeded.`)

    return parseJsonResponse<T>(text)
  }

  throw new Error(
    [
      `No supported Gemini model endpoint was found for this API key. Tried: ${failures.join(', ')}`,
      'Verify Generative Language API access is enabled for this key and that model access is available in your region.',
    ].join(' '),
  )
}

async function fetchPageMarkdown(url: string) {
  const readerUrl = `https://r.jina.ai/${url}`
  const response = await fetch(readerUrl)

  if (!response.ok) {
    throw new Error(`Jina Reader request failed (${response.status}).`)
  }

  const markdown = await response.text()
  if (!markdown.trim()) {
    throw new Error('Jina Reader returned empty content.')
  }

  return markdown
}

function stageMessage(stage: ProcessingStage) {
  if (stage === 'reading') {
    return 'Reading your landing page...'
  }

  if (stage === 'analyzing') {
    return 'Analyzing your ad creative...'
  }

  if (stage === 'personalizing') {
    return 'Personalizing for message-match...'
  }

  return ''
}

function markdownPreview(markdown: string, limit = 1600) {
  const trimmed = markdown.trim()

  if (!trimmed) {
    return ''
  }

  if (trimmed.length <= limit) {
    return trimmed
  }

  return `${trimmed.slice(0, limit)}\n\n...`
}

function fallbackHtml(markdown: string) {
  const safe = markdown
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/\n/g, '<br/>')

  return `
  <!doctype html>
  <html>
    <head>
      <meta charset="UTF-8" />
      <meta name="viewport" content="width=device-width, initial-scale=1.0" />
      <title>Personalized Landing Page</title>
      <style>
        body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; background: #ffffff; color: #0f172a; margin: 0; padding: 32px; line-height: 1.6; }
        .wrap { max-width: 960px; margin: 0 auto; }
        h1 { margin-top: 0; }
        .note { margin-bottom: 20px; padding: 12px 14px; background: #eef2ff; border: 1px solid #c7d2fe; border-radius: 10px; }
      </style>
    </head>
    <body>
      <div class="wrap">
        <div class="note">Gemini did not return valid HTML, so this is a text fallback.</div>
        <div>${safe}</div>
      </div>
    </body>
  </html>
  `
}

export default function App() {
  const [apiKey, setApiKey] = useState('')
  const [adImage, setAdImage] = useState<File | null>(null)
  const [adImagePreview, setAdImagePreview] = useState('')
  const [landingUrl, setLandingUrl] = useState('')
  const [stage, setStage] = useState<ProcessingStage>('idle')
  const [error, setError] = useState('')
  const [pageMarkdown, setPageMarkdown] = useState('')
  const [adSignals, setAdSignals] = useState<AdSignals | null>(null)
  const [output, setOutput] = useState<PersonalizationOutput | null>(null)
  const [copied, setCopied] = useState(false)
  const [stepLogs, setStepLogs] = useState<StepLog[]>([])

  function appendLog(message: string, level: StepLog['level'] = 'info') {
    setStepLogs((current) => [
      ...current,
      {
        id: Date.now() + current.length,
        time: new Date().toLocaleTimeString(),
        message,
        level,
      },
    ])
  }

  useEffect(() => {
    const storedKey = window.localStorage.getItem(apiKeyStorageKey)
    if (storedKey) {
      setApiKey(storedKey)
    }
  }, [])

  useEffect(() => {
    if (!adImage) {
      setAdImagePreview('')
      return
    }

    const objectUrl = URL.createObjectURL(adImage)
    setAdImagePreview(objectUrl)

    return () => URL.revokeObjectURL(objectUrl)
  }, [adImage])

  const loading = stage === 'reading' || stage === 'analyzing' || stage === 'personalizing'
  const analysisReady = Boolean(pageMarkdown.trim() && adSignals)

  const renderedHtml = useMemo(() => {
    if (!output?.html?.trim()) {
      return pageMarkdown ? fallbackHtml(pageMarkdown) : ''
    }

    return output.html
  }, [output, pageMarkdown])

  async function handleAnalyzeInputs() {
    setError('')
    setCopied(false)
    setStepLogs([])
    appendLog('Analyze flow started.')

    const normalizedUrl = normalizeUrl(landingUrl)
    const key = apiKey.trim() || import.meta.env.VITE_GEMINI_API_KEY?.trim() || ''

    if (!key) {
      setError('Enter your Gemini API key in settings to continue.')
      appendLog('Missing Gemini API key.', 'error')
      return
    }

    if (!adImage) {
      setError('Upload an ad creative image first.')
      appendLog('No ad image uploaded.', 'error')
      return
    }

    if (!normalizedUrl) {
      setError('Enter a valid landing page URL.')
      appendLog('Landing page URL is invalid.', 'error')
      return
    }

    window.localStorage.setItem(apiKeyStorageKey, key)

    try {
      setStage('reading')
      setOutput(null)
      appendLog('Reading landing page markdown from Jina Reader...')

      const fetchedMarkdown = await fetchPageMarkdown(normalizedUrl)
      setPageMarkdown(fetchedMarkdown)
      appendLog('Landing page markdown extracted successfully.', 'success')

      setStage('analyzing')
      appendLog('Analyzing ad creative with Gemini Vision...')

      const adSignalsPrompt = `Analyze this ad creative and extract:
- headline
- offer or hook
- target audience signal
- tone (urgent / aspirational / trust-based / playful etc.)
- primary CTA

Return ONLY a JSON object with keys: headline, offer, audience, tone, cta`

      const signals = await callGemini<AdSignals>({
        apiKey: key,
        prompt: adSignalsPrompt,
        imageFile: adImage,
        onLog: (message) => appendLog(message),
      })

      setAdSignals(signals)
      appendLog('Ad analysis completed.', 'success')
      setStage('done')
    } catch (caughtError) {
      const message = caughtError instanceof Error ? caughtError.message : 'Something went wrong.'
      setError(message)
      appendLog(`Analyze failed: ${message}`, 'error')
      setStage('idle')
    }
  }

  async function handlePersonalizePage() {
    setError('')
    setCopied(false)

    const key = apiKey.trim() || import.meta.env.VITE_GEMINI_API_KEY?.trim() || ''

    if (!key) {
      setError('Enter your Gemini API key in settings to continue.')
      appendLog('Missing Gemini API key.', 'error')
      return
    }

    if (!analysisReady || !adSignals) {
      setError('Run Analyze Inputs first so you can preview analysis before personalizing.')
      appendLog('Personalization blocked until analysis is complete.', 'error')
      return
    }

    window.localStorage.setItem(apiKeyStorageKey, key)

    try {
      setStage('personalizing')
      appendLog('Personalization started.')
      appendLog('Sending page markdown + ad signals to Gemini...')

      const personalizationPrompt = `You are a senior CRO strategist.

Here is a landing page content in markdown:
${pageMarkdown}

Here are the signals from an ad creative that points to this page:
${JSON.stringify(adSignals, null, 2)}

Your job:
1. Rewrite the landing page copy to maximize message-match with the ad
2. Keep the exact same page structure, sections, and layout
3. Only change: hero headline, subheadline, CTA text, value proposition bullets, and any offer-related body copy
4. Match the tone and terminology of the ad
5. Reconstruct the full page as clean, styled HTML with inline CSS
6. Make it visually clean - use a white background, modern sans-serif font, good spacing
7. Preserve the logical flow and section order of the original page

Also return a personalization_rationale - 2-3 lines explaining what you changed and why.

Return a JSON object with keys:
- html (the full reconstructed personalized page as a string)
- personalization_rationale (string)
- changes (array of objects with keys: element, original, rewritten)`

      const result = await callGemini<PersonalizationOutput>({
        apiKey: key,
        prompt: personalizationPrompt,
        onLog: (message) => appendLog(message),
      })

      setOutput({
        html: result.html ?? '',
        personalization_rationale:
          result.personalization_rationale ??
          'Copy was adapted to align ad promise and landing page narrative for stronger message-match.',
        changes: Array.isArray(result.changes) ? result.changes : [],
      })
      appendLog('Personalized page generated successfully.', 'success')
      setStage('done')
    } catch (caughtError) {
      const message = caughtError instanceof Error ? caughtError.message : 'Something went wrong.'
      setError(message)
      appendLog(`Personalization failed: ${message}`, 'error')
      setStage('idle')
    }
  }

  async function handleCopyHtml() {
    if (!renderedHtml) {
      return
    }

    try {
      await navigator.clipboard.writeText(renderedHtml)
      setCopied(true)
      appendLog('Copied personalized HTML to clipboard.', 'success')
      window.setTimeout(() => setCopied(false), 1500)
    } catch {
      setError('Unable to copy HTML. Please copy from preview manually.')
      appendLog('Copy to clipboard failed.', 'error')
    }
  }

  return (
    <main className="app">
      <header className="topbar">
        <div className="brand-block">
          <p className="brand">MessageMatch</p>
          <p className="tagline">Personalize any landing page to match your ad. Instantly.</p>
        </div>

        <div className="topbar-actions">
          <details className="how-it-works">
            <summary>How it works</summary>
            <ol>
              <li>Reads your page with Jina Reader.</li>
              <li>Analyzes your ad with Gemini Vision.</li>
              <li>Rebuilds your page copy for message-match.</li>
            </ol>
          </details>

          <label className="api-field" htmlFor="gemini-api-key">
            <span>Gemini API key</span>
            <input
              id="gemini-api-key"
              type="password"
              value={apiKey}
              onChange={(event) => setApiKey(event.target.value)}
              placeholder="Paste your Gemini API key"
            />
          </label>
        </div>
      </header>

      <section className="input-shell">
        <div className="input-grid">
          <section className="card">
            <h2>Ad Creative</h2>
            <p className="muted">Upload the actual ad image (Meta, Google, LinkedIn, etc.)</p>
            <label className="upload" htmlFor="ad-file">
              <span>Upload ad image</span>
              <input
                id="ad-file"
                type="file"
                accept="image/*"
                onChange={(event) => setAdImage(event.target.files?.[0] ?? null)}
              />
            </label>
            {adImagePreview ? <img src={adImagePreview} alt="Ad creative preview" className="preview-image" /> : null}
          </section>

          <section className="card">
            <h2>Landing Page URL</h2>
            <p className="muted">Enter the destination page URL you want to personalize.</p>
            <input
              className="url-input"
              type="url"
              value={landingUrl}
              onChange={(event) => setLandingUrl(event.target.value)}
              placeholder="https://example.com"
            />

            <div className="action-stack">
              <button type="button" className="cta" disabled={loading} onClick={handleAnalyzeInputs}>
                {stage === 'reading' || stage === 'analyzing' ? 'Analyzing...' : '1) Analyze Inputs'}
              </button>

              <button
                type="button"
                className="cta cta-secondary"
                disabled={loading || !analysisReady}
                onClick={handlePersonalizePage}
              >
                {stage === 'personalizing' ? 'Personalizing...' : '2) Personalize Page'}
              </button>
            </div>

            <p className="hint">Analyze first to preview both extraction steps, then click personalize.</p>

            {loading ? <p className="status">{stageMessage(stage)}</p> : null}
            {error ? <p className="error">{error}</p> : null}
          </section>

          <section className={`card step-card ${adSignals ? 'step-ready' : ''}`}>
            <h3>Step 1: Ad Analysis</h3>
            {adSignals ? (
              <dl className="signals-grid">
                <div>
                  <dt>Headline</dt>
                  <dd>{adSignals.headline || '-'}</dd>
                </div>
                <div>
                  <dt>Offer / Hook</dt>
                  <dd>{adSignals.offer || '-'}</dd>
                </div>
                <div>
                  <dt>Audience</dt>
                  <dd>{adSignals.audience || '-'}</dd>
                </div>
                <div>
                  <dt>Tone</dt>
                  <dd>{adSignals.tone || '-'}</dd>
                </div>
                <div>
                  <dt>CTA</dt>
                  <dd>{adSignals.cta || '-'}</dd>
                </div>
              </dl>
            ) : (
              <p className="muted">Run Analyze Inputs to see extracted ad signals.</p>
            )}
          </section>

          <section className={`card step-card ${pageMarkdown ? 'step-ready' : ''}`}>
            <h3>Step 2: Extracted Landing Page Content</h3>
            {pageMarkdown ? (
              <pre className="markdown-preview">{markdownPreview(pageMarkdown)}</pre>
            ) : (
              <p className="muted">Run Analyze Inputs to preview extracted landing page markdown.</p>
            )}
          </section>

          <section className="card step-card log-card">
            <h3>Step Logs</h3>
            {stepLogs.length > 0 ? (
              <ul className="log-list">
                {stepLogs.map((log) => (
                  <li key={log.id} className={`log-item log-${log.level}`}>
                    <span className="log-time">{log.time}</span>
                    <span className="log-message">{log.message}</span>
                  </li>
                ))}
              </ul>
            ) : (
              <p className="muted">Run Analyze Inputs to start collecting step-by-step logs.</p>
            )}
          </section>
        </div>
      </section>

      <section className={`output-shell ${output ? 'visible' : ''}`}>
        <section className="card insight-card">
          <div className="card-headline">
            <span className="icon" aria-hidden="true">
              i
            </span>
            <h3>Why this works</h3>
          </div>
          <p>
            {output?.personalization_rationale ??
              'Generate a personalized page to see the CRO rationale behind the message-match updates.'}
          </p>
          {adSignals ? (
            <div className="signals">
              <span>Detected ad signals:</span>
              <p>
                {adSignals.headline} | {adSignals.offer} | {adSignals.tone}
              </p>
            </div>
          ) : null}
        </section>

        <section className="card">
          <h3>What Changed</h3>
          {output?.changes?.length ? (
            <div className="diff-wrap">
              <table className="diff-table">
                <thead>
                  <tr>
                    <th>Element</th>
                    <th>Original Copy</th>
                    <th>Personalized Copy</th>
                  </tr>
                </thead>
                <tbody>
                  {output.changes.map((change, index) => (
                    <tr key={`${change.element}-${index}`}>
                      <td>{change.element || 'Copy block'}</td>
                      <td>{change.original || '-'}</td>
                      <td className="personalized">{change.rewritten || '-'}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : (
            <p className="muted">Your copy diff will appear here after generation.</p>
          )}
        </section>

        <section className="card">
          <div className="preview-head">
            <h3>Your Personalized Landing Page</h3>
            <button className="copy-button" type="button" onClick={handleCopyHtml} disabled={!renderedHtml}>
              {copied ? 'Copied' : 'Copy HTML'}
            </button>
          </div>
          <div className="iframe-shell">
            {renderedHtml ? (
              <iframe title="Personalized landing page" srcDoc={renderedHtml} sandbox="" />
            ) : (
              <div className="preview-empty">Your personalized page preview will render here.</div>
            )}
          </div>
        </section>
      </section>
    </main>
  )
}
