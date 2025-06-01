// API Configuration
// 環境変数から取得、ローカル開発時は直接値を設定可能

export const API_CONFIG = {
  // Jina AI API Configuration
  JINA_API_TOKEN: process.env.JINA_API_TOKEN || """",
  
  // Tavily Search API Configuration
  TAVILY_API_KEY: process.env.TAVILY_API_KEY || """",
  
  // Google Custom Search API Configuration
  GOOGLE_API_KEY: process.env.GOOGLE_API_KEY || """",
  GOOGLE_SEARCH_ENGINE_ID: process.env.GOOGLE_SEARCH_ENGINE_ID || """"
};