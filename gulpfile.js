'use strict';

const gulp = require('gulp');
const build = require('@microsoft/sp-build-web');
const webpack = require('webpack');
build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

build.configureWebpack.mergeConfig({
    additionalConfiguration: (generatedConfiguration) => {
      generatedConfiguration.devtool = 'source-map';
  
      for (var i = 0; i < generatedConfiguration.plugins.length; i++) {
        const plugin = generatedConfiguration.plugins[i];
        if (plugin instanceof webpack.optimize.UglifyJsPlugin) {
          plugin.options.sourceMap = true;
          break;
        }
      }
  
      return generatedConfiguration;
    }
  });

build.initialize(gulp);
