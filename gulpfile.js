'use strict';

const build = require('@microsoft/sp-build-web');
const eslint = require('gulp-eslint'); // for ts-eslint
const buildtypescript = require('@microsoft/gulp-core-build-typescript');  // for ts-eslint
const log = require('fancy-log');
const TerserPlugin = require('terser-webpack-plugin-legacy');

/** Terser "ESBUILD" minifier */
const ESBUILD_TerserOptions = {
  minify: TerserPlugin.esbuildMinify,
  // `terserOptions` options will be passed to `esbuild` -  Link to options - https://esbuild.github.io/api/#minify  
  terserOptions: {},
}

// use installed TypeScript version
const typeScriptConfig = require('@microsoft/gulp-core-build-typescript/lib/TypeScriptConfiguration');
typeScriptConfig.TypeScriptConfiguration.setTypescriptCompiler(require('typescript'));

// ensure the installed react version is used and replace UglifyJS minifier with Terser
build.configureWebpack.mergeConfig({
  additionalConfiguration: (generatedConfiguration) => {
    // use installed React version
    generatedConfiguration.externals = generatedConfiguration.externals.filter(
      (name) => !['react', 'react-dom'].includes(name)
    );
    // Look for the minifier plugin (currently: UglifyJS) in the plugin list and replace it with a Terser plugin
    generatedConfiguration.plugins.forEach((plugin, i) => {
      if (plugin.options && plugin.options.mangle) {
        // use the Terser plugin with the "ESBUILD" minifier instead of the standard minifier for maximized size reduction
        const tersePlugin = new TerserPlugin(ESBUILD_TerserOptions);
        log.info(`Replace webpack minifier [${plugin.constructor.name}] with [${tersePlugin.constructor.name}]`);
        generatedConfiguration.plugins.splice(i, 1, tersePlugin);
      }
    });
    return generatedConfiguration;
  }
});

// use typescript-eslint instead of tslint
buildtypescript.tslint.enabled = false; // disable tslint
const eslintSubTask = build.subTask('eslint', function (gulp, buildOptions, done) {
  return (
    gulp
      .src(['src/**/*.{ts,tsx}'])
      // eslint() attaches the lint output to the "eslint" property of the file object so it can be used by other modules.
      .pipe(eslint())
      // eslint.format() outputs the lint results to the console. Alternatively use eslint.formatEach() (see Docs).
      .pipe(eslint.format())
      // To have the process exit with an error code (1) on lint error, return the stream and pipe to failAfterError last.
      .pipe(eslint.failAfterError())
  );
});
build.rig.addPreBuildTask(build.task('eslint-task', eslintSubTask)); // for ts-eslint

// avoid gulp warnings due to sass files
build.addSuppression(/Warning - \[sass\] The local CSS class/gi);


build.initialize(require('gulp'));
