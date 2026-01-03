import { fileURLToPath } from "node:url"
import { includeIgnoreFile } from "@eslint/compat"
import js from "@eslint/js"
import stylistic from "@stylistic/eslint-plugin"
import { defineConfig } from "eslint/config"
import globals from "globals"
import ts from "typescript-eslint"

const gitignorePath = fileURLToPath(new URL("./.gitignore", import.meta.url))
const parserOptions = {
	projectService: {
		allowDefaultProject: ["*.js", "vitest.config.ts", "tests/*.ts"]
	},
	parser: ts.parser
}

export default defineConfig(
	includeIgnoreFile(gitignorePath),
	js.configs.recommended,
	...ts.configs.recommended,
	{
		plugins: {
			"@stylistic": stylistic
		},
		languageOptions: {
			globals: {
				...globals.node,
				SpreadsheetApp: "readonly",
				ScriptApp: "readonly",
				Logger: "readonly"
			}
		},
		rules: {
			"no-console": ["warn", { allow: ["warn", "error"] }],
			"prefer-const": "warn",
			"no-var": "error",
			"semi": ["warn", "never"],
			"no-trailing-spaces": "warn",
			"quotes": ["warn", "double", { avoidEscape: true }],
			"@stylistic/member-delimiter-style": [
				"warn",
				{
					multiline: {
						delimiter: "none"
					},
					singleline: {
						delimiter: "comma",
						requireLast: false
					}
				}
			]
		}
	},
	{
		files: ["**/*.ts", "**/*.tsx"],
		ignores: ["tests/**/*.ts"],
		languageOptions: {
			parserOptions
		}
	},
	{
		files: ["tests/**/*.ts"],
		languageOptions: {
			parser: ts.parser
		}
	}
)
