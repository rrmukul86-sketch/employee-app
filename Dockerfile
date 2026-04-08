# Use Node.js as the base image for both frontend build and backend runtime
FROM node:22-slim AS base
WORKDIR /app

# Stage 1: Build the React/Vite Frontend
FROM base AS builder
COPY package*.json ./
RUN npm install
COPY . .
RUN npm run build

# Stage 2: Final Production Runtime (Unified Backend/Frontend)
FROM base AS runner
# Install only production dependencies
COPY package*.json ./
RUN npm install --production

# Copy built frontend assets
COPY --from=builder /app/dist ./dist

# Copy backend scripts
COPY scripts/ ./scripts/
COPY .env ./

# Expose ports for Vite (3000) and Upload Service (3001)
EXPOSE 3000 3001

# Command to run the service
# Note: For production, we'll run the upload-service which handles attachments
CMD ["node", "scripts/upload-service.mjs"]
