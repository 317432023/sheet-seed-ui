const { defineConfig } = require('@vue/cli-service')
module.exports = defineConfig({
  transpileDependencies: true,
  lintOnSave: true,
  publicPath: '/',
  // webpack-dev-server 相关配置
  devServer: {
    open: true,
    host: 'localhost',
    port: 8000,
    https: false,
    // http 代理配置
    proxy: {
      '/server': { // 1、'/server' 本身代表http//localhost:${port}/server
        target: 'http://127.0.0.1:8080/server', // 2、target将http//localhost:${port}/server指向'http://127.0.0.1:8080/server'
        changeOrigin: true,
        pathRewrite: {
          '^/server': '' // 3、pathRewrite: {'^/server' : ''}又把http//localhost:${port}/server后面的server去除了
        }
      },
    },
  },
})
