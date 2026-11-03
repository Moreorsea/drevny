// https://nuxt.com/docs/api/configuration/nuxt-config
import { fileURLToPath } from 'node:url'

const baseURL = process.env.NUXT_APP_BASE_URL || '/'

export default defineNuxtConfig({
  modules: [
    '@nuxt/eslint',
    '@nuxt/ui'
  ],

  // SPA: conversion runs in the browser → works on GitHub Pages
  ssr: false,

  devtools: {
    enabled: true
  },

  css: ['~/assets/css/main.css'],

  app: {
    baseURL
  },

  alias: {
    '#isolation': fileURLToPath(new URL('./shared/isolation', import.meta.url))
  },

  nitro: {
    preset: 'static'
  },

  vite: {
    optimizeDeps: {
      include: ['exceljs', 'docx']
    },
    define: {
      global: 'globalThis'
    }
  },

  compatibilityDate: '2026-06-30',

  eslint: {
    config: {
      stylistic: {
        commaDangle: 'never',
        braceStyle: '1tbs'
      }
    }
  }
})
