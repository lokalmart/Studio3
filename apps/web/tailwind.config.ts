import type { Config } from 'tailwindcss';

const config: Config = {
  content: ['./app/**/*.{ts,tsx}', './components/**/*.{ts,tsx}', './lib/**/*.{ts,tsx}'],
  theme: {
    extend: {
      colors: {
        lokalmart: {
          bg: '#0f0b08',
          card: '#1a1511',
          line: '#3a3028',
          amber: '#f59e0b',
          orange: '#fb923c'
        }
      }
    }
  },
  plugins: []
};
export default config;
