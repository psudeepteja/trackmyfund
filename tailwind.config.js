/** @type {import('tailwindcss').Config} */
module.exports = {
  content: [
    './pages/**/*.{js,ts,jsx,tsx,mdx}',
    './components/**/*.{js,ts,jsx,tsx,mdx}',
    './app/**/*.{js,ts,jsx,tsx,mdx}',
  ],
  theme: {
    extend: {
      fontFamily: {
        display: ['Playfair Display', 'Georgia', 'serif'],
        body: ['DM Sans', 'sans-serif'],
        mono: ['DM Mono', 'monospace'],
      },
      colors: {
        ink: '#0D0D0D',
        paper: '#F5F0E8',
        gold: '#C9A84C',
        'gold-light': '#F0D898',
        'gold-dark': '#8B6914',
        emerald: '#1A5C3A',
        'emerald-light': '#2D8A58',
        cream: '#FDF8F0',
        muted: '#8A8070',
      },
    },
  },
  plugins: [],
}
