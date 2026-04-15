FROM node:22-alpine
WORKDIR /app
RUN mkdir -p public
COPY GoSafe_Dashboard_Startco2026.html public/index.html
COPY server.mjs mailerlite.js gosafe_vc_outreach_email.html ./
ENV NODE_ENV=production
ENV PORT=8080
EXPOSE 8080
USER node
CMD ["node", "server.mjs"]
