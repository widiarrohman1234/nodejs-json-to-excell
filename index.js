const json2xls = require("json2xls");
const fs = require("fs");
const readline = require("readline");

// Membuat interface pembacaan baris untuk membaca dari terminal
const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout,
});

// Fungsi untuk mengubah JSON menjadi Excel
const convertJsonToExcel = (jsonArray, excelFilePath) => {
  const xls = json2xls(jsonArray); // Menggunakan format data yang sesuai
  fs.writeFileSync(excelFilePath, xls, "binary");
  console.log("File Excel berhasil dibuat:", excelFilePath);
  rl.close(); // Menutup interface pembacaan baris setelah selesai
};

// Meminta alamat file JSON dari pengguna
rl.question("Masukkan alamat file JSON: ", (jsonFilePath) => {
  try {
    // Membaca data JSON dari file
    const jsonData = JSON.parse(fs.readFileSync(jsonFilePath, "utf8"));

    // Nama file Excel yang akan dihasilkan
    const excelFilePath = "ouput.xlsx";

    // Panggil fungsi untuk mengubah JSON menjadi Excel
    convertJsonToExcel(jsonData.data, excelFilePath); // Menggunakan array data di dalam objek
  } catch (error) {
    console.error("Terjadi kesalahan:", error.message);
  }
});
