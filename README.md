# MessageMatch

AI-powered ad-to-landing-page personalization tool.

Requires Node 18.18 or newer.

## Setup

Set `VITE_GEMINI_API_KEY` in your environment or paste a key in the app settings.

## Scripts

- `npm run dev` starts the development server.
- `npm run build` creates a production build.
- `npm run preview` previews the built app locally.

## Flow

1. Upload an ad creative image.
2. Paste a landing page URL.
3. The app reads the page via Jina Reader, extracts ad signals with Gemini Vision, then reconstructs personalized HTML with Gemini.
4. It shows rationale, a copy diff table, and a rendered personalized preview with copy-to-clipboard HTML.