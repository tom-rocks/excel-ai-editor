/** @type {import('tailwindcss').Config} */
export default {
  content: [
    "./index.html",
    "./src/**/*.{js,ts,jsx,tsx}",
  ],
  theme: {
    extend: {
      colors: {
        'midnight': '#0a0a0f',
        'surface': '#12121a',
        'surface-light': '#1a1a24',
        'accent': '#00d9ff',
        'accent-dim': '#00a3bf',
        'success': '#00ff88',
        'warning': '#ffaa00',
      },
      fontFamily: {
        'display': ['Outfit', 'sans-serif'],
        'mono': ['JetBrains Mono', 'monospace'],
      },
    },
  },
  plugins: [],
}
