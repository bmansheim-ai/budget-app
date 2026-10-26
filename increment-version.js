const fs = require('fs');
const path = require('path');

const packageJsonPath = path.join(__dirname, 'package.json');
const packageJson = JSON.parse(fs.readFileSync(packageJsonPath, 'utf8'));

// Split the version string (e.g., "1.0.0" -> [1, 0, 0])
const versionParts = packageJson.version.split('.');

// Increment the patch version number (the last digit)
versionParts[2] = parseInt(versionParts[2], 10) + 1;
packageJson.version = versionParts.join('.');

// Save the updated package.json back to disk
fs.writeFileSync(packageJsonPath, JSON.stringify(packageJson, null, 2) + '\n');
console.log(`[Build System] Automatically incremented version to: ${packageJson.version}`);