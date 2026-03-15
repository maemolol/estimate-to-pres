/** @type {import('tailwindcss').Config} */
export default {
  content: ['./index.html', './src/**/*.{js,ts,jsx,tsx}'],
  theme: {
    extend: {
      fontFamily: {
        sans: ['DM Sans', 'sans-serif'],
        display: ['Fraunces', 'serif'],
        mono: ['DM Mono', 'monospace'],
      },
      colors: {
        ink: { DEFAULT: '#1a1a18', 2: '#5a5a56', 3: '#9a9a94' },
        surface: { DEFAULT: '#fafaf7', card: '#ffffff', alt: '#f4f4f0' },
        brand: {
          green: '#1a5c3a', 'green-mid': '#2d8a58', 'green-lt': '#e8f5ee',
          amber: '#b85c00', 'amber-lt': '#fff3e6',
          navy: '#1a3a5c', 'navy-lt': '#e6f0f8',
        },
        border: { DEFAULT: '#e8e8e0', strong: '#d0d0c8' },
      }
    }
  },
  plugins: []
}
