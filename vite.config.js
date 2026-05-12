import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// ⚠️ base 값을 본인의 GitHub 레포지토리 이름으로 변경하세요
// 예: 레포 이름이 'asset-management' 이면 '/asset-management/'
export default defineConfig({
  plugins: [react()],
  base: '/asset-management/',
})
