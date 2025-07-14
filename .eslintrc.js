/* eslint-env node */
/* rules: 
  - eslint: https://eslint.org/docs/latest/rules/  
  - typescript: https://typescript-eslint.io/rules/
  - react: https://github.com/jsx-eslint/eslint-plugin-react/blob/HEAD/docs/rules/no-array-index-key.md
*/
module.exports = {
    extends: ['eslint:recommended', 'plugin:@typescript-eslint/recommended', 'plugin:react/recommended'],
    parser: '@typescript-eslint/parser',
    plugins: ['@typescript-eslint'],
    root: true,
    "parserOptions": {
      "ecmaFeatures": {
        "jsx": true
      },
      "ecmaVersion": 12,
      "sourceType": "module"
    },
    "rules": {      
      "no-var": "error",      
      "@typescript-eslint/no-explicit-any": "warn",
      "@typescript-eslint/no-var-requires": "warn",
      "@typescript-eslint/no-shadow": "warn",
      "@typescript-eslint/default-param-last": "warn",
      "@typescript-eslint/semi": "warn",
      "react/destructuring-assignment": "warn",
      "react/no-unused-prop-types": "warn",
      "react/no-unknown-property": "warn",
      "react/no-array-index-key": "warn",
      "react/jsx-no-useless-fragment": "warn",
      "react/no-this-in-sfc": "error",
      "react/no-unused-state": "error",
      "react/jsx-pascal-case": "error",
      "react/jsx-no-duplicate-props": "error"
    },
    "ignorePatterns": [
      "*.js",
      "**/coverage",
      "**/dist",
      "**/etc",
      "**/lib",
      "**/sharepoint",
      "**/lib-amd",
      "**/lib-commonjs",
      "**/node_modules",
      "**/temp",
      "**/*.scss.ts"
    ],
    "settings": {
      "react": {
        "version": "detect"
      }
    }    
  };
