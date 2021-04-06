import typescript from '@rollup/plugin-typescript';
import resolve from '@rollup/plugin-node-resolve'

const config = {
  input: 'src/index.ts',
  output: {
    dir: 'output',
    format: 'esm',
    sourcemap: true
  },
  plugins: [
    resolve({modulesOnly: true}),
    typescript({
      sourceMap: true
    })
  ]
};

export default config;