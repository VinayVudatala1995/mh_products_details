const axios = require("axios");
const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");
const download = require("image-downloader");
require("dotenv").config();

// Path to the existing Excel file
const filePath = "./list.xlsx";
const outputPath = "./file_with_image_paths.xlsx";

// Initialize Excel workbook
const workbook = new ExcelJS.Workbook();

// Set the correct path for the images folder relative to your script location
const downloadFolder = path.join(__dirname, "images");
// Ensure the download folder exists
if (!fs.existsSync(downloadFolder)) {
  fs.mkdirSync(downloadFolder, { recursive: true });
}

async function readExcel() {
  await workbook.xlsx.readFile(filePath);
  const worksheet = workbook.getWorksheet(1);

  // Define headers and ensure columns are correctly assigned
  const headers = [,];
  worksheet.columns = headers.map((title, index) => ({
    header: title,
    key: title,
    width: 20,
  }));

  worksheet.eachRow(async (row, rowNumber) => {
    if (rowNumber > 1) {
      // Assuming the first row is headers
      const productDesc = row.getCell("Product_Desc").value;
      const packaging = row.getCell("Packing").value;
      const query = productDesc + " " + packaging;
      const productData = await fetchProductData(query, productDesc);

      if (productData) {
        row.getCell("Product Title").value =
          productData.title || "No Data Found";
        row.getCell("Product Description").value =
          productData.description || "No Data Found";
        if (productData.imageUrl) {
          const savedImagePath = await downloadImage(productData.imageUrl);
          row.getCell("Image_Path").value = savedImagePath || "No Data Found";
        }
      } else {
        row.getCell("Product Title").value = "No Data Found";
        row.getCell("Product Description").value = "No Data Found";
        row.getCell("Image_Path").value = "No Data Found";
      }
    }
  });

  await workbook.xlsx.writeFile(outputPath);
}

async function readExcel() {
  await workbook.xlsx.readFile(filePath);
  const worksheet = workbook.getWorksheet(1);

  // Define headers if they are not already present
  if (worksheet.getRow(1).getCell(1).value !== "Product_Desc") {
    const headers = [
      "Product_Desc",
      "Packing",
      "Mfg_Code",
      "PTR",
      "MRP",
      "Tax_Per",
      "Product Title",
      "Product Description",
      "Image_Path",
    ];
    headers.forEach((title, index) => {
      worksheet.getRow(1).getCell(index + 1).value = title;
    });
  }

  const rowsToUpdate = [];

  // Collect all row update promises
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) {
      const productDesc = `${row.getCell(1).value} ${row.getCell(2).value}`;
      rowsToUpdate.push(
        updateRowWithData(row, productDesc, row.getCell(1).value)
      );
    }
  });

  // Wait for all rows to be updated
  await Promise.all(rowsToUpdate);

  // Save the updated workbook
  await workbook.xlsx.writeFile(outputPath);
}

async function updateRowWithData(row, productDesc, title) {
  const productData = await fetchProductData(productDesc, title);
  if (productData) {
    row.getCell(7).value = productData.title || "No Data Found";
    row.getCell(8).value = productData.description || "No Data Found";
    if (productData.imageUrl) {
      const savedImagePath = await downloadImage(productData.imageUrl);
      row.getCell(9).value = savedImagePath || "No Data Found";
    }
  } else {
    row.getCell(7).value = "No Data Found";
    row.getCell(8).value = "No Data Found";
    row.getCell(9).value = "No Data Found";
  }
}

async function fetchProductData(query, productDesc) {
  console.log(process.env.key);
  const options = {
    method: "GET",
    url: "https://real-time-amazon-data.p.rapidapi.com/search",
    params: {
      query: query,
      page: "1",
      country: "IN",
      sort_by: "RELEVANCE",
      product_condition: "NEW",
      is_prime: "false",
    },
    headers: {
      "x-rapidapi-key": process.env.YOUR_KEY,
      "x-rapidapi-host": "real-time-amazon-data.p.rapidapi.com",
    },
  };

  try {
    const response = await axios.request(options);
    console.log(`options :: ${options}`);
    console.log(`productDesc :: ${productDesc}`);
    console.log(`Response :: ${response}`);
    const queryWords = productDesc.toLowerCase().split(" ");

    // Find the product that contains all words from the query in its title
    const exactMatch =
      response.data.data.products.length > 1
        ? response.data.data.products.find((product) => {
            // Normalize and split the title into words
            const product_title = product.product_title
              .toLowerCase()
              .split(" ");
            // Check if every word from the query is included in the title
            // return queryWords.every((word) => product_title.includes(word));
            var val = false;
            for (let index = 0; index < queryWords.length; index++) {
              const element = queryWords[index];
              if (product_title.includes(element)) {
                val = true;
              } else {
                val = false;
                break;
              }
            }
            return val;
          })
        : response.data.data.length == 1
        ? response.data.data.products[1]
        : undefined;
    console.log(`exactMatch :: ${exactMatch}`);
    return exactMatch
      ? {
          title: exactMatch.product_title,
          description: exactMatch.product_url,
          imageUrl: exactMatch.product_photo,
        }
      : null;
  } catch (error) {
    console.error(error);
    return null;
  }
}

async function downloadImage(url) {
  const options = {
    url: url,
    dest: path.join(downloadFolder, path.basename(url)), // Ensure the file is saved in your 'images' folder
    timeout: 15000, // Example timeout setting
  };

  try {
    const { filename } = await download.image(options);
    return path.relative(".", filename);
  } catch (error) {
    console.error("Error downloading image:", error);
    return null;
  }
}

readExcel();
