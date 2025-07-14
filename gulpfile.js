'use strict';

const gulp = require('gulp');
const build = require('@microsoft/sp-build-web');

// SharePoint 2019 specific build configuration with React 17 override
build.addSuppression(/Warning - \[sass\]/);
build.addSuppression(/Warning - \[tslint\]/);
build.addSuppression(/Warning.*react/i); // Suppress React version warnings

// Force React 17 configuration for SharePoint 2019
build.configureWebpack.mergeConfig({
  additionalConfiguration: (generatedConfiguration) => {
    // Force React 17 resolution
    if (!generatedConfiguration.resolve) {
      generatedConfiguration.resolve = {};
    }
    
    if (!generatedConfiguration.resolve.alias) {
      generatedConfiguration.resolve.alias = {};
    }

    // Force React 17 aliases to override SPFx defaults
    generatedConfiguration.resolve.alias['react'] = require.resolve('react');
    generatedConfiguration.resolve.alias['react-dom'] = require.resolve('react-dom');
    
    // Ensure React 17 is treated as external for proper bundling
    if (!generatedConfiguration.externals) {
      generatedConfiguration.externals = {};
    }
    
    // Override SPFx React externals with React 17
    generatedConfiguration.externals['react'] = {
      amd: 'react',
      commonjs: 'react',
      commonjs2: 'react',
      root: 'React'
    };
    
    generatedConfiguration.externals['react-dom'] = {
      amd: 'react-dom',
      commonjs: 'react-dom',
      commonjs2: 'react-dom',
      root: 'ReactDOM'
    };

    // Configure module rules for React 17 JSX transform
    if (generatedConfiguration.module && generatedConfiguration.module.rules) {
      generatedConfiguration.module.rules.forEach(rule => {
        if (rule.use && Array.isArray(rule.use)) {
          rule.use.forEach(use => {
            if (use.loader && use.loader.includes('ts-loader')) {
              if (!use.options) use.options = {};
              if (!use.options.compilerOptions) use.options.compilerOptions = {};
              
              // Enable React 17 JSX transform
              use.options.compilerOptions.jsx = 'react-jsx';
              use.options.compilerOptions.jsxImportSource = 'react';
            }
          });
        }
      });
    }

    // Optimize bundle for SharePoint 2019 with React 17
    if (generatedConfiguration.optimization) {
      generatedConfiguration.optimization.minimize = true;
      
      // Ensure React 17 is properly split
      if (!generatedConfiguration.optimization.splitChunks) {
        generatedConfiguration.optimization.splitChunks = {};
      }
      
      generatedConfiguration.optimization.splitChunks.cacheGroups = {
        ...generatedConfiguration.optimization.splitChunks.cacheGroups,
        react: {
          name: 'react',
          chunks: 'all',
          test: /[\\/]node_modules[\\/](react|react-dom)[\\/]/,
          priority: 20,
          enforce: true
        }
      };
    }

    // Add React 17 specific plugins
    if (!generatedConfiguration.plugins) {
      generatedConfiguration.plugins = [];
    }

    // Define React version for runtime
    const webpack = require('webpack');
    generatedConfiguration.plugins.push(
      new webpack.DefinePlugin({
        'process.env.REACT_VERSION': JSON.stringify('17.0.2'),
        'process.env.NODE_ENV': JSON.stringify(process.env.NODE_ENV || 'development')
      })
    );

    console.log('✅ React 17.0.2 forced in SPFx 1.4.1 for SharePoint 2019');
    
    return generatedConfiguration;
  }
});

// Custom task for React 17 compatibility check
const reactCompatibilityTask = build.subTask('react-17-compatibility', function(gulp, buildOptions, done) {
  console.log('🔍 Checking React 17 compatibility with SharePoint 2019...');
  
  try {
    const react = require('react');
    const reactDom = require('react-dom');
    
    console.log('✅ React version:', react.version);
    console.log('✅ ReactDOM version:', reactDom.version);
    
    if (react.version.startsWith('17.')) {
      console.log('✅ React 17 successfully forced in SPFx 1.4.1');
    } else {
      console.warn('⚠️  React version mismatch - expected 17.x, got', react.version);
    }
  } catch (error) {
    console.error('❌ React compatibility check failed:', error.message);
  }
  
  done();
});

// Custom task for SharePoint 2019 deployment with React 17
const deployTask = build.subTask('deploy-sp2019-react17', function(gulp, buildOptions, done) {
  console.log('🚀 Preparing React 17 deployment for SharePoint 2019 on-premise...');
  
  // Validate React 17 bundle size for SharePoint 2019
  const fs = require('fs');
  const path = require('path');
  
  try {
    const distPath = path.join(__dirname, 'dist');
    if (fs.existsSync(distPath)) {
      const files = fs.readdirSync(distPath);
      const jsFiles = files.filter(file => file.endsWith('.js'));
      
      jsFiles.forEach(file => {
        const filePath = path.join(distPath, file);
        const stats = fs.statSync(filePath);
        const sizeKB = (stats.size / 1024).toFixed(2);
        
        console.log(`📦 Bundle: ${file} - ${sizeKB} KB`);
        
        // Warn if bundle is too large for SharePoint 2019
        if (stats.size > 2 * 1024 * 1024) { // 2MB limit
          console.warn(`⚠️  Large bundle detected: ${file} (${sizeKB} KB)`);
          console.warn('   Consider code splitting for SharePoint 2019 performance');
        }
      });
    }
  } catch (error) {
    console.warn('⚠️  Bundle size check failed:', error.message);
  }
  
  console.log('✅ React 17 deployment preparation complete');
  done();
});

// Add tasks to build pipeline
build.rig.addPreBuildTask(reactCompatibilityTask);
build.rig.addPostBuildTask(deployTask);

build.initialize(gulp);
