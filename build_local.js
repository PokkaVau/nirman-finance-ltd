const fs = require('fs');

// Read modular templates
const templateHtml = fs.readFileSync('template.html', 'utf8');
const stylesHtml = fs.readFileSync('Styles.html', 'utf8');
const scriptsHtml = fs.readFileSync('Scripts.html', 'utf8');

// Replace GAS template scriptlets with inlined CSS & JS
let bundled = templateHtml.replace("<?!= include('Styles'); ?>", stylesHtml);
bundled = bundled.replace("<?!= include('Scripts'); ?>", scriptsHtml);

// Output compiled standalone files
fs.writeFileSync('index.html', bundled, 'utf8');
fs.writeFileSync('index_local.html', bundled, 'utf8');

console.log('Successfully bundled index.html and index_local.html: ' + bundled.length + ' bytes');
