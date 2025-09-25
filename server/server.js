import express from "express";
import cors from "cors";
import fs from "fs";
import path from "path";
import { exec } from "child_process";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());

app.get("/", (req, res) => {
  res.send("Web Office Server - Templates API");
});

// Send the list of templates
app.get("/templates", (req, res) => {
  const templatesDir = path.join(__dirname, "templates");
  
  fs.readdir(templatesDir, (err, files) => {
    if (err) {
      console.error("Error reading templates directory:", err);
      return res.status(500).json({ error: "Failed to read templates directory" });
    }
    
    // Filter for document files and remove extensions for cleaner API
    const templateFiles = files
      .filter(file => {
        const ext = path.extname(file).toLowerCase();
        return ['.docx', '.doc', '.odt', '.rtf', '.txt'].includes(ext);
      })
      .map(file => ({
        name: path.basename(file, path.extname(file)),
        filename: file,
        extension: path.extname(file)
      }));
    
    res.json({
      templates: templateFiles,
      count: templateFiles.length
    });
  });
});

// Serve static assets (images, css, etc.)
app.use('/assets', express.static(path.join(__dirname, 'temp', 'assets')));

// Convert selected template to HTML using LibreOffice with enhanced formatting
app.get("/template/:name", async (req, res) => {
  const templateName = req.params.name;
  const templatesDir = path.join(__dirname, "templates");
  const outputDir = path.join(__dirname, "temp");
  const assetsDir = path.join(outputDir, "assets");
  
  try {
    // Create temp and assets directories if they don't exist
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }
    if (!fs.existsSync(assetsDir)) {
      fs.mkdirSync(assetsDir, { recursive: true });
    }
    
    // Find the template file
    const files = fs.readdirSync(templatesDir);
    const templateFile = files.find(file => 
      path.basename(file, path.extname(file)) === templateName
    );
    
    if (!templateFile) {
      return res.status(404).json({ error: "Template not found" });
    }
    
    const inputPath = path.join(templatesDir, templateFile);
    const outputFileName = templateName + ".html";
    const outputPath = path.join(outputDir, outputFileName);
    
    console.log(`Converting document: ${inputPath}`);
    console.log(`Output directory: ${outputDir}`);
    console.log(`Assets directory: ${assetsDir}`);
    
    // Try multiple LibreOffice conversion methods for better image extraction
    
    // Method 1: Convert with image embedding
    const embedCommand = `libreoffice --headless --convert-to "html:HTML (StarWriter):EmbedImages" --outdir "${outputDir}" "${inputPath}"`;
    
    exec(embedCommand, (embedError, embedStdout, embedStderr) => {
      console.log("Embed conversion stdout:", embedStdout);
      console.log("Embed conversion stderr:", embedStderr);
      
      if (embedError) {
        console.log("Embedded image conversion failed, trying standard HTML...");
        
        // Method 2: Standard HTML conversion
        const standardCommand = `libreoffice --headless --convert-to html --outdir "${outputDir}" "${inputPath}"`;
        
        exec(standardCommand, (stdError, stdStdout, stdStderr) => {
          console.log("Standard conversion stdout:", stdStdout);
          console.log("Standard conversion stderr:", stdStderr);
          
          if (stdError) {
            console.error("Standard HTML conversion failed:", stdError);
            return res.status(500).json({ 
              error: "Failed to convert document to HTML",
              details: stdError.message 
            });
          }
          
          // Extract images using multiple methods
          extractImagesMultipleMethods(inputPath, assetsDir, templateName, outputDir, () => {
            processConvertedHTML(outputPath, res, templateName, assetsDir, outputDir, PORT);
          });
        });
        return;
      }
      
      console.log("Embedded conversion successful, extracting images...");
      // Extract images using multiple methods
      extractImagesMultipleMethods(inputPath, assetsDir, templateName, outputDir, () => {
        processConvertedHTML(outputPath, res, templateName, assetsDir, outputDir, PORT);
      });
    });
    
  } catch (error) {
    console.error("Template conversion error:", error);
    res.status(500).json({ error: "Internal server error" });
  }
});

// Function to extract images using multiple methods
function extractImagesMultipleMethods(inputPath, assetsDir, templateName, outputDir, callback) {
  console.log("Starting multiple image extraction methods...");
  
  let methodsCompleted = 0;
  const totalMethods = 3;
  
  const checkCompletion = () => {
    methodsCompleted++;
    if (methodsCompleted >= totalMethods) {
      console.log("All image extraction methods completed");
      callback();
    }
  };
  
  // Method 1: Extract from ODT conversion
  extractFromODT(inputPath, assetsDir, templateName, checkCompletion);
  
  // Method 2: Extract from DOCX structure (if it's a DOCX file)
  extractFromDOCX(inputPath, assetsDir, templateName, checkCompletion);
  
  // Method 3: Check LibreOffice output directory for generated images
  setTimeout(() => {
    extractFromOutputDir(outputDir, assetsDir, checkCompletion);
  }, 1000); // Wait a bit for LibreOffice to finish
}

// Method 1: Extract images from ODT conversion
function extractFromODT(inputPath, assetsDir, templateName, callback) {
  console.log("Method 1: Extracting images via ODT conversion...");
  
  const tempDir = path.join(assetsDir, '..', 'temp_extract');
  
  if (!fs.existsSync(tempDir)) {
    fs.mkdirSync(tempDir, { recursive: true });
  }
  
  const odtPath = path.join(tempDir, `${templateName}.odt`);
  const extractCommand = `libreoffice --headless --convert-to odt --outdir "${tempDir}" "${inputPath}"`;
  
  exec(extractCommand, (error, stdout, stderr) => {
    console.log("ODT conversion output:", stdout);
    console.log("ODT conversion errors:", stderr);
    
    if (!error && fs.existsSync(odtPath)) {
      // ODT files are ZIP archives, extract images
      const unzipCommand = `cd "${tempDir}" && unzip -j "${odtPath}" Pictures/* -d "${assetsDir}" 2>/dev/null || true`;
      
      exec(unzipCommand, (unzipError, unzipStdout) => {
        console.log("Unzip output:", unzipStdout);
        
        // Clean up temp ODT file
        if (fs.existsSync(odtPath)) {
          fs.unlinkSync(odtPath);
        }
        
        callback();
      });
    } else {
      console.log("ODT conversion failed or file not found");
      callback();
    }
  });
}

// Method 2: Extract images from DOCX structure
function extractFromDOCX(inputPath, assetsDir, templateName, callback) {
  console.log("Method 2: Extracting images from DOCX structure...");
  
  const ext = path.extname(inputPath).toLowerCase();
  if (ext !== '.docx') {
    console.log("Not a DOCX file, skipping DOCX extraction");
    callback();
    return;
  }
  
  const tempDir = path.join(assetsDir, '..', 'temp_docx');
  
  if (!fs.existsSync(tempDir)) {
    fs.mkdirSync(tempDir, { recursive: true });
  }
  
  // DOCX files are ZIP archives, extract media folder
  const unzipCommand = `cd "${tempDir}" && unzip -j "${inputPath}" word/media/* -d "${assetsDir}" 2>/dev/null || true`;
  
  exec(unzipCommand, (error, stdout, stderr) => {
    console.log("DOCX media extraction output:", stdout);
    if (stderr) console.log("DOCX media extraction errors:", stderr);
    
    callback();
  });
}

// Method 3: Extract images from LibreOffice output directory
function extractFromOutputDir(outputDir, assetsDir, callback) {
  console.log("Method 3: Checking LibreOffice output directory for images...");
  
  try {
    const outputFiles = fs.readdirSync(outputDir);
    console.log("Files in output directory:", outputFiles);
    
    const imageFiles = outputFiles.filter(file => {
      const ext = path.extname(file).toLowerCase();
      return ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.svg', '.webp'].includes(ext);
    });
    
    console.log("Found image files in output:", imageFiles);
    
    imageFiles.forEach(imageFile => {
      const sourcePath = path.join(outputDir, imageFile);
      const destPath = path.join(assetsDir, imageFile);
      
      try {
        fs.copyFileSync(sourcePath, destPath);
        console.log(`Copied image: ${imageFile} to assets`);
      } catch (copyError) {
        console.error("Error copying image:", copyError);
      }
    });
    
  } catch (dirError) {
    console.log("Error reading output directory:", dirError);
  }
  
  callback();
}

// Function to process and enhance converted HTML
function processConvertedHTML(outputPath, res, templateName, assetsDir, outputDir, port) {
  fs.readFile(outputPath, "utf8", (readErr, htmlContent) => {
    if (readErr) {
      console.error("Error reading converted HTML:", readErr);
      return res.status(500).json({ error: "Failed to read converted HTML" });
    }
    
    try {
      console.log("Processing HTML for template:", templateName);
      
      // Enhance HTML content with better formatting and styling
      let enhancedHTML = enhanceHTMLFormatting(htmlContent, templateName);
      
      // Extract and handle embedded images
      enhancedHTML = handleEmbeddedImages(enhancedHTML, assetsDir, templateName, port);
      
      // Handle LibreOffice generated image files
      enhancedHTML = handleLibreOfficeImages(enhancedHTML, assetsDir, templateName, outputDir, port);
      
      // Clean up the temporary HTML file
      fs.unlink(outputPath, (unlinkErr) => {
        if (unlinkErr) {
          console.error("Error cleaning up temp file:", unlinkErr);
        }
      });
      
      // Send the enhanced HTML content
      res.setHeader('Content-Type', 'text/html; charset=utf-8');
      res.send(enhancedHTML);
      
    } catch (processError) {
      console.error("Error processing HTML:", processError);
      res.status(500).json({ error: "Failed to process converted HTML" });
    }
  });
}

// Function to enhance HTML formatting
function enhanceHTMLFormatting(htmlContent, templateName) {
  // Add enhanced CSS for better formatting
  const enhancedCSS = `
    <style>
      body {
        font-family: 'Times New Roman', Times, serif;
        line-height: 1.6;
        margin: 20px;
        background-color: #ffffff;
      }
      
      /* Preserve text formatting */
      .bold, strong, b { font-weight: bold !important; }
      .italic, em, i { font-style: italic !important; }
      .underline, u { text-decoration: underline !important; }
      
      /* Text alignment */
      .text-left { text-align: left !important; }
      .text-center { text-align: center !important; }
      .text-right { text-align: right !important; }
      .text-justify { text-align: justify !important; }
      
      /* Preserve colors */
      [style*="color"] { color: inherit !important; }
      [style*="background"] { background: inherit !important; }
      
      /* Tables */
      table {
        border-collapse: collapse;
        width: 100%;
        margin: 10px 0;
      }
      
      table, th, td {
        border: 1px solid #ddd;
      }
      
      th, td {
        padding: 8px;
        text-align: left;
      }
      
      /* Images */
      img {
        max-width: 100%;
        height: auto;
        display: block;
        margin: 10px 0;
      }
      
      /* Headings */
      h1, h2, h3, h4, h5, h6 {
        margin: 20px 0 10px 0;
        font-weight: bold;
      }
      
      /* Lists */
      ul, ol {
        margin: 10px 0;
        padding-left: 30px;
      }
      
      /* Paragraphs */
      p {
        margin: 10px 0;
      }
      
      /* Preserve LibreOffice styles */
      [class*="western"] { font-family: inherit; }
      [class*="ctl"] { font-family: inherit; }
      [class*="cjk"] { font-family: inherit; }
    </style>
  `;
  
  // Insert enhanced CSS into the HTML head
  if (htmlContent.includes('<head>')) {
    htmlContent = htmlContent.replace('<head>', '<head>' + enhancedCSS);
  } else if (htmlContent.includes('<html>')) {
    htmlContent = htmlContent.replace('<html>', '<html><head>' + enhancedCSS + '</head>');
  } else {
    htmlContent = enhancedCSS + htmlContent;
  }
  
  // Add viewport meta tag for responsive design
  const viewportMeta = '<meta name="viewport" content="width=device-width, initial-scale=1.0">';
  if (htmlContent.includes('<head>')) {
    htmlContent = htmlContent.replace('<head>', '<head>' + viewportMeta);
  }
  
  return htmlContent;
}

// Function to handle LibreOffice generated image files
function handleLibreOfficeImages(htmlContent, assetsDir, templateName, outputDir, port) {
  console.log("Checking for LibreOffice generated images in:", outputDir);
  
  // Look for image files that LibreOffice might have generated
  try {
    const outputFiles = fs.readdirSync(outputDir);
    const imageFiles = outputFiles.filter(file => {
      const ext = path.extname(file).toLowerCase();
      return ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.svg'].includes(ext);
    });
    
    console.log("Found image files:", imageFiles);
    
    imageFiles.forEach(imageFile => {
      const sourcePath = path.join(outputDir, imageFile);
      const destPath = path.join(assetsDir, imageFile);
      
      // Move image to assets directory
      try {
        fs.copyFileSync(sourcePath, destPath);
        fs.unlinkSync(sourcePath); // Clean up original
        
        // Update HTML to reference the moved image
        const oldSrc = imageFile;
        const newSrc = `http://localhost:${port}/assets/${imageFile}`;
        htmlContent = htmlContent.replace(new RegExp(oldSrc, 'g'), newSrc);
        
        console.log(`Moved image: ${imageFile} -> ${newSrc}`);
      } catch (moveError) {
        console.error("Error moving image file:", moveError);
      }
    });
  } catch (dirError) {
    console.error("Error reading output directory:", dirError);
  }
  
  return htmlContent;
}

// Function to handle embedded images
function handleEmbeddedImages(htmlContent, assetsDir, templateName, port) {
  console.log("Processing embedded images for template:", templateName);
  
  // List all images found in assets directory
  try {
    const assetImages = fs.readdirSync(assetsDir);
    console.log("Images available in assets directory:", assetImages);
  } catch (dirError) {
    console.log("Assets directory not accessible:", dirError);
  }
  
  // Look for embedded images in the HTML and extract them
  const imgRegex = /<img[^>]+src="([^"]+)"[^>]*>/gi;
  let match;
  let imageCounter = 0;
  let updatedHTML = htmlContent;
  
  // Reset regex index
  imgRegex.lastIndex = 0;
  
  while ((match = imgRegex.exec(htmlContent)) !== null) {
    const fullImgTag = match[0];
    const imgSrc = match[1];
    console.log("Found image source:", imgSrc.length > 100 ? imgSrc.substring(0, 100) + "..." : imgSrc);
    
    // Handle base64 embedded images
    if (imgSrc.startsWith('data:image/')) {
      try {
        const mimeMatch = imgSrc.match(/data:image\/([^;]+);base64,(.+)/);
        if (mimeMatch) {
          const imageType = mimeMatch[1];
          const base64Data = mimeMatch[2];
          const imageName = `${templateName}_embedded_${imageCounter}.${imageType}`;
          const imagePath = path.join(assetsDir, imageName);
          
          // Save the image
          fs.writeFileSync(imagePath, base64Data, 'base64');
          
          // Replace the base64 src with the served file path
          const newSrc = `http://localhost:${port}/assets/${imageName}`;
          const newImgTag = fullImgTag.replace(imgSrc, newSrc);
          updatedHTML = updatedHTML.replace(fullImgTag, newImgTag);
          
          console.log(`✅ Extracted embedded image: ${imageName}`);
          imageCounter++;
        }
      } catch (imgError) {
        console.error("❌ Error processing embedded image:", imgError);
      }
    }
    // Handle relative image paths and file references
    else if (!imgSrc.startsWith('http') && !imgSrc.startsWith('/assets/')) {
      const imageName = path.basename(imgSrc);
      
      // Check if the image exists in assets directory
      const assetPath = path.join(assetsDir, imageName);
      if (fs.existsSync(assetPath)) {
        const newSrc = `http://localhost:${port}/assets/${imageName}`;
        const newImgTag = fullImgTag.replace(imgSrc, newSrc);
        updatedHTML = updatedHTML.replace(fullImgTag, newImgTag);
        console.log(`✅ Updated image path: ${imgSrc} -> ${newSrc}`);
      } else {
        console.log(`⚠️  Image not found in assets: ${imageName}`);
        
        // Try to find similar named files
        try {
          const assetFiles = fs.readdirSync(assetsDir);
          const similarFile = assetFiles.find(file => 
            file.toLowerCase().includes(imageName.toLowerCase().split('.')[0])
          );
          
          if (similarFile) {
            const newSrc = `http://localhost:${port}/assets/${similarFile}`;
            const newImgTag = fullImgTag.replace(imgSrc, newSrc);
            updatedHTML = updatedHTML.replace(fullImgTag, newImgTag);
            console.log(`✅ Matched similar image: ${imgSrc} -> ${newSrc}`);
          }
        } catch (searchError) {
          console.log("Could not search for similar images:", searchError);
        }
      }
    }
  }
  
  // Final check: list all image tags in the processed HTML
  const finalImgRegex = /<img[^>]+src="([^"]+)"[^>]*>/gi;
  const finalMatches = [...updatedHTML.matchAll(finalImgRegex)];
  console.log("Final image sources in HTML:");
  finalMatches.forEach((match, index) => {
    console.log(`  ${index + 1}. ${match[1]}`);
  });
  
  return updatedHTML;
}

app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
});
