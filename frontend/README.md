# Attendance Tracker - Next.js Frontend

Modern web-based GUI for the Attendance Tracker using Next.js, React, and Tailwind CSS.

## Features

- 🎨 Modern UI with Tailwind CSS
- ⚡ Fast and responsive with Next.js
- 🤖 Integrated with Gemini API
- 📊 Real-time roster information
- 🔄 Live output updates

## Prerequisites

- Node.js 18+ and npm
- Python Flask backend running on `http://localhost:5001`
- Gemini API key set in environment variable `GEMINI_API_KEY`

## Installation

1. Navigate to the frontend directory:
```bash
cd frontend
```

2. Install dependencies:
```bash
npm install
```

## Running the Application

1. Make sure the Flask backend is running:
```bash
# In the project root
python app.py
```

2. Start the Next.js development server:
```bash
npm run dev
```

3. Open your browser to `http://localhost:3000`

## Project Structure

```
frontend/
├── app/
│   ├── layout.tsx       # Root layout
│   ├── page.tsx         # Main page
│   └── globals.css      # Global styles
├── components/
│   ├── StatusBar.tsx    # Top status bar
│   ├── RosterInfo.tsx   # Roster information panel
│   ├── ActionPanel.tsx  # Action buttons panel
│   └── OutputPanel.tsx  # Output display panel
├── package.json
├── next.config.js
├── tailwind.config.js
└── tsconfig.json
```

## API Endpoints

The frontend communicates with the Flask backend via these endpoints:

- `POST /api/roster/load` - Load roster file
- `GET /api/roster/info` - Get roster information
- `POST /api/attendance/process` - Process attendance with Gemini
- `POST /api/query` - Natural language query
- `POST /api/student/find` - Find student information
- `POST /api/dsl/execute` - Execute DSL code

## Development

- The app uses Next.js 14 with App Router
- Styling is done with Tailwind CSS
- TypeScript is configured for type safety
- CORS is enabled on the Flask backend for `http://localhost:3000`

## Building for Production

```bash
npm run build
npm start
```

