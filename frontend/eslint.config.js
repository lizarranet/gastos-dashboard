import js from '@eslint/js'
import globals from 'globals'
import reactHooks from 'eslint-plugin-react-hooks'
import reactRefresh from 'eslint-plugin-react-refresh'
import { defineConfig, globalIgnores } from 'eslint/config'

export default defineConfig([
  globalIgnores(['dist']),
  {
    files: ['**/*.{js,jsx}'],
    extends: [
      js.configs.recommended,
      reactHooks.configs.flat.recommended,
      reactRefresh.configs.vite,
    ],
    languageOptions: {
      globals: globals.browser,
      parserOptions: { ecmaFeatures: { jsx: true } },
    },
    rules: {
      // Estas llamadas sincronizan el componente con la API; React admite este patrón
      // siempre que las respuestas se limpien o cancelen al desmontar.
      'react-hooks/set-state-in-effect': 'off',
      // El generador visual se conserva para una reactivación próxima de su botón.
      'no-unused-vars': ['error', { varsIgnorePattern: '^buildAnalisisIAPrompt$' }],
    },
  },
])
