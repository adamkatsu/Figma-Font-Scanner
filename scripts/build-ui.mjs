#!/usr/bin/env node
/**
 * Bundles ui/src/main.ts with esbuild and inlines CSS + JS into root ui.html
 * from ui/index.html template.
 */
import * as esbuild from 'esbuild';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.join(__dirname, '..');
const watch = process.argv.includes('--watch');

async function buildOnce() {
  const result = await esbuild.build({
    entryPoints: [path.join(root, 'ui/src/main.ts')],
    bundle: true,
    write: false,
    format: 'iife',
    target: ['es2017'],
    logLevel: 'info'
  });

  const js = result.outputFiles[0].text;
  const css = fs.readFileSync(path.join(root, 'ui/src/styles.css'), 'utf8');
  let html = fs.readFileSync(path.join(root, 'ui/index.html'), 'utf8');

  html = html.replace('<!--INJECT_CSS-->', `<style>\n${css}\n</style>`);
  html = html.replace('<!--INJECT_JS-->', `<script>\n${js}\n</script>`);

  fs.writeFileSync(path.join(root, 'ui.html'), html, 'utf8');
  console.log('Wrote ui.html');
}

if (watch) {
  const ctx = await esbuild.context({
    entryPoints: [path.join(root, 'ui/src/main.ts')],
    bundle: true,
    write: false,
    format: 'iife',
    target: ['es2017'],
    logLevel: 'info',
    plugins: [
      {
        name: 'inline-html',
        setup(build) {
          build.onEnd(async result => {
            if (result.errors.length) return;
            const js = result.outputFiles[0].text;
            const css = fs.readFileSync(path.join(root, 'ui/src/styles.css'), 'utf8');
            let html = fs.readFileSync(path.join(root, 'ui/index.html'), 'utf8');
            html = html.replace('<!--INJECT_CSS-->', `<style>\n${css}\n</style>`);
            html = html.replace('<!--INJECT_JS-->', `<script>\n${js}\n</script>`);
            fs.writeFileSync(path.join(root, 'ui.html'), html, 'utf8');
            console.log('Wrote ui.html');
          });
        }
      }
    ]
  });
  await ctx.watch();
  // Also rebuild when CSS or index.html changes
  fs.watch(path.join(root, 'ui/src/styles.css'), () => ctx.rebuild());
  fs.watch(path.join(root, 'ui/index.html'), () => ctx.rebuild());
  console.log('Watching UI…');
} else {
  await buildOnce();
}
