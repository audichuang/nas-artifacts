import { serve } from '@hono/node-server'
import { createAppServer } from './app-server.ts'
import { ArtifactStore } from './artifact-store.ts'
import { loadConfig } from './config.ts'
import { createContentServer } from './content-server.ts'
import { createPublicAppServer, publicContentConfig } from './public-server.ts'

const config = loadConfig()
const store = new ArtifactStore(config.dataDir)
const host = process.env.LISTEN_HOST ?? '0.0.0.0'

const servers = [
  serve({ fetch: createAppServer(store, config).fetch, port: config.appPort, hostname: host }),
  serve({ fetch: createContentServer(store, config).fetch, port: config.contentPort, hostname: host })
]
console.log(`artifacts: app ${config.appOrigin} (:${config.appPort}), content ${config.contentOrigin} (:${config.contentPort}), data ${config.dataDir}`)
if (config.public) {
  servers.push(
    serve({ fetch: createPublicAppServer(store, config).fetch, port: config.public.appPort, hostname: host }),
    serve({ fetch: createContentServer(store, publicContentConfig(config), { publicOnly: true }).fetch, port: config.public.contentPort, hostname: host })
  )
  console.log(`artifacts: public app ${config.public.appOrigin} (:${config.public.appPort}), public content ${config.public.contentOrigin} (:${config.public.contentPort})`)
}

function shutdown(): void {
  let open = servers.length
  for (const server of servers) {
    server.close(() => {
      open -= 1
      if (open === 0) {
        store.close()
        process.exit(0)
      }
    })
  }
  setTimeout(() => process.exit(0), 5000).unref()
}
process.on('SIGTERM', shutdown)
process.on('SIGINT', shutdown)
