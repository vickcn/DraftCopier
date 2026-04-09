/** @type {import('next').NextConfig} */
const nextConfig = {
  async rewrites() {
    // 只有在本地開發環境時，才把 /api 導向本機的 FastAPI (6311 port)
    // 部署到 Vercel 時 (production) 會自動套用 Vercel 內建的路由，我們不干涉。
    if (process.env.NODE_ENV === 'development') {
      return [
        {
          source: '/api/:path*',
          destination: 'http://localhost:6311/api/:path*', 
        },
      ]
    }
    return []
  },
}

module.exports = nextConfig
