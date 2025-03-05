const fs = require('fs');
const path = require('path');
const { logger } = require('../src/utils/logger');

const cleanDirs = ['dist', '.webpack'];

function removeDir(dirPath) {
  try {
    if (fs.existsSync(dirPath)) {
      const files = fs.readdirSync(dirPath);
      for (const file of files) {
        const filePath = path.join(dirPath, file);
        if (fs.statSync(filePath).isDirectory()) {
          removeDir(filePath);
        } else {
          fs.unlinkSync(filePath);
        }
      }
      fs.rmdirSync(dirPath);
      logger.info(`Removed directory: ${dirPath}`);
    }
  } catch (error) {
    logger.error(`Error cleaning directory ${dirPath}:`, error);
  }
}

for (const dir of cleanDirs) {
  const dirPath = path.resolve(__dirname, '..', dir);
  removeDir(dirPath);
}

logger.info('Clean completed successfully.');
