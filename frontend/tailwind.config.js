/** @type {import('tailwindcss').Config} */
export default {
  content: [
    "./index.html",
    "./src/**/*.{js,ts,jsx,tsx}",
  ],
  theme: {
    extend: {
      colors: {
        // Sidebar
        'sidebar-bg': '#1A1A2E',
        'sidebar-hover': '#252542',
        'sidebar-active': '#2D2D4A',

        // Main area
        'main-bg': '#F5F1EB',
        'card-bg': '#FFFFFF',

        // Accents
        'primary': '#4A7C59',
        'primary-hover': '#3D6B4A',
        'accent-yellow': '#F5D547',
        'accent-pink': '#E8B4B8',
        'accent-olive': '#8B956D',
        'accent-purple': '#7C5295',

        // Status
        'success': '#4CAF50',
        'success-bg': '#E8F5E9',
        'error': '#EF5350',
        'error-bg': '#FFEBEE',

        // Text
        'text-primary': '#1A1A2E',
        'text-secondary': '#5A5A7A',
        'text-muted': '#9A9ABF',
      },
    },
  },
  plugins: [],
}
