import { fileURLToPath } from 'node:url'
import { includeIgnoreFile } from '@eslint/compat'
import js from '@eslint/js'
import stylistic from '@stylistic/eslint-plugin'
import { defineConfig } from 'eslint/config'
import globals from 'globals'
import ts from 'typescript-eslint'

const gitignorePath = fileURLToPath(new URL('./.gitignore', import.meta.url))
const parserOptions = {
	projectService: {
		allowDefaultProject: ['*.js', 'vitest.config.ts']
	},
	parser: ts.parser
}

export default defineConfig(
	includeIgnoreFile(gitignorePath),
	js.configs.recommended,
	...ts.configs.recommended,
	{
		plugins: {
			'@stylistic': stylistic
		},
		languageOptions: {
			globals: {
				...globals.node,
				SpreadsheetApp: 'readonly',
				ScriptApp: 'readonly',
				Logger: 'readonly'
			}
		},
		rules: {
			'no-undef': 'off',
			'@typescript-eslint/no-unused-vars': 'off',
			'no-console': ['warn', { allow: ['warn', 'error'] }],
			'prefer-const': 'warn',
			'no-var': 'error',
			'semi': ['warn', 'never'],
			'no-trailing-spaces': 'warn',
			'@stylistic/member-delimiter-style': [
				'warn',
				{
					multiline: {
						delimiter: 'none'
					},
					singleline: {
						delimiter: 'comma',
						requireLast: false
					}
				}
			]
		}
	},
	{
		files: ['**/*.ts', '**/*.tsx'],
		languageOptions: {
			parserOptions
		}
	}
)
