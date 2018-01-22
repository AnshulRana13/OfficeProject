var compressor = require('node-minify');
compressor.minify({
  compressor: 'gcc',
  input: 'PkgInfoCreateExcel.js',
  output: 'bar.js',
  callback: function (err, min) {}
});