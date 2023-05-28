import { nodeResolve } from '@rollup/plugin-node-resolve'
import ts from 'rollup-plugin-ts'

const extensions = ['.ts', '.js']

const preventTreeShakingPlugin = () => {
  return {
    name: 'no-treeshaking',
    resolveId(id, importer) {
      if (!importer) {
        // let's not theeshake entry points, as we're not exporting anything in Apps Script files
        return { id, moduleSideEffects: 'no-treeshake' }
      }
      return null
    }
  }
}

export default {
  input: './index.ts',
  output: {
    dir: 'build',
    format: 'esm'
  },
  plugins: [
    preventTreeShakingPlugin(),
    nodeResolve({ extensions }),
    ts({
      tsconfig: {
        fileName: '../../tsconfig.json',
        hook: (resolvedConfig) => ({ ...resolvedConfig, downlevelIteration: true })
      }
    })
  ]
}
