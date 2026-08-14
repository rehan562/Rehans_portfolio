const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
const output = path.join(root, 'build');

fs.rmSync(output, { recursive: true, force: true });
fs.mkdirSync(output, { recursive: true });

for (const file of ['index.html', 'styles.css']) {
  fs.copyFileSync(
    path.join(root, file),
    path.join(output, file)
  );
}

fs.cpSync(
  path.join(root, 'assets'),
  path.join(output, 'assets'),
  { recursive: true }
);