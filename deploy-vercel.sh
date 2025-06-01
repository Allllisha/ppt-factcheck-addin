#!/bin/bash

# Vercel deployment script
echo "Deploying to Vercel..."

# Deploy with Vercel CLI
vercel --prod \
  --env JINA_API_TOKEN="$JINA_API_TOKEN" \
  --env TAVILY_API_KEY="$TAVILY_API_KEY" \
  --env GOOGLE_API_KEY="$GOOGLE_API_KEY" \
  --env GOOGLE_SEARCH_ENGINE_ID="$GOOGLE_SEARCH_ENGINE_ID" \
  --yes

echo "Deployment complete!"