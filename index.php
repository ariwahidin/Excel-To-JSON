<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <meta http-equiv="X-UA-Compatible" content="ie=edge" />
    <title>Excel to Table</title>
    <script type="text/javascript" src="js/xlsx.full.min.js"></script>
    <style>
        table {
            border-collapse: collapse;
            width: 100%;
        }

        th,
        td {
            border: 1px solid #dddddd;
            text-align: left;
            padding: 8px;
        }

        th {
            background-color: #f2f2f2;
        }

        .total-qty-cell {
            text-align: right;
        }
    </style>
</head>

<body>
    <input type="file" id="fileUpload" accept=".xls,.xlsx" /><br />
    <button type="button" id="uploadExcel">Convert</button>
    <pre id="jsonData"></pre>
    <div id="divTable"></div>

    <script>
        var selectedFile;

        // Objek yang menampung nama kolom yang ditentukan dari sistem
        var systemColumnNames = {
            columnName1: "item_code",
            columnName2: "location",
            columnName3: "expire", // Kolom ketiga harus berisi tanggal
            columnName4: "qty",
            // Tambahkan lebih banyak nama kolom jika diperlukan
        };

        document.getElementById("fileUpload").addEventListener("change", function(event) {
            selectedFile = event.target.files[0];
        });

        document.getElementById("uploadExcel").addEventListener("click", function() {
            if (selectedFile) {
                var fileReader = new FileReader();
                fileReader.onload = function(event) {
                    var data = event.target.result;

                    try {
                        var workbook = XLSX.read(data, {
                            type: "binary",
                        });

                        var excelData = {
                            rows: [],
                            totalQty: 0
                        };

                        workbook.SheetNames.forEach((sheet) => {
                            let table = document.createElement("table");
                            table.setAttribute("id", "excelTable");
                            let rowObject = XLSX.utils.sheet_to_row_object_array(
                                workbook.Sheets[sheet]
                            );
                            var headers = Object.keys(rowObject[0]);

                            // Validasi nama kolom
                            var systemColumnKeys = Object.keys(systemColumnNames);
                            if (headers.length !== systemColumnKeys.length ||
                                !systemColumnKeys.every(key => headers.includes(systemColumnNames[key]))) {

                                alert("Column names do not match the system-defined column names.");
                                return;
                            }

                            // Validasi kolom ketiga untuk format tanggal
                            var dateRegex = /^\d{4}-\d{2}-\d{2}$/; // Format tanggal: YYYY-MM-DD
                            var invalidDateFound = false;
                            for (var i = 0; i < rowObject.length; i++) {
                                var rowData = rowObject[i];
                                if (!dateRegex.test(convertExcelSerialToDate(rowData[systemColumnNames.columnName3]))) {
                                    invalidDateFound = true;
                                    break;
                                }
                            }

                            if (invalidDateFound) {
                                alert("Data in column 'Expire' must be in date format (YYYY-MM-DD).");
                                return; // Stop further processing if invalid date found
                            }

                            var headerRow = table.insertRow(-1);
                            headers.forEach(function(header) {
                                var cell = headerRow.insertCell(-1);
                                cell.innerHTML = header;
                            });



                            // Konversi nilai tanggal dari format serial Excel ke format JavaScript
                            rowObject.forEach(function(rowData) {
                                var dateSerial = parseFloat(rowData[systemColumnNames.columnName3]);
                                var convertedDate = convertExcelSerialToDate(dateSerial);
                                rowData[systemColumnNames.columnName3] = convertedDate;
                                excelData.rows.push(rowData);

                                // Hitung total kuantitas
                                var qty = parseInt(rowData[systemColumnNames.columnName4]) || 0;
                                excelData.totalQty += qty;
                            });
                        });

                        // Tampilkan tabel
                        displayTable(excelData);

                        console.log(excelData);
                    } catch (error) {
                        console.error("Error parsing Excel file:", error);
                        alert("Error parsing Excel file. Please make sure the file is valid.");
                    }
                };
                fileReader.readAsBinaryString(selectedFile);
            }
        });



        // Function to display table
        function displayTable(excelData) {
            var table = document.createElement("table");
            table.setAttribute("id", "excelTable");

            var headers = Object.values(systemColumnNames); // Mengambil nilai kolom dari objek systemColumnNames

            // Header row
            var headerRow = table.insertRow(-1);
            headers.unshift("No.");
            headers.forEach(function(header) {
                var cell = headerRow.insertCell(-1);
                cell.innerHTML = header;
            });

            // Data rows
            excelData.rows.forEach(function(rowData, index) {
                var row = table.insertRow(-1);
                headers.forEach(function(header) {
                    var cell = row.insertCell(-1);
                    if (header === "No.") {
                        cell.innerHTML = index + 1; // Nomor urut dimulai dari 1
                    } else {
                        cell.innerHTML = rowData[header];
                    }
                });
            });

            // Total Qty row
            var totalRow = table.insertRow(-1);
            var totalQtyCell = totalRow.insertCell(-1);
            totalQtyCell.colSpan = headers.length - 1; // Mengatur kolom keempat
            totalQtyCell.style.fontWeight = "bold";
            totalQtyCell.innerHTML = "Total Qty:";
            var totalValueCell = totalRow.insertCell(-1);
            totalQtyCell.className = "total-qty-cell"; // Menambahkan kelas untuk penataan CSS
            totalValueCell.style.fontWeight = "bold";
            totalValueCell.innerHTML = excelData.totalQty;

            var divTable = document.getElementById("divTable");
            divTable.innerHTML = "";
            divTable.appendChild(table);
        }

        // Function to convert Excel serial date to JavaScript date
        function convertExcelSerialToDate(serial) {
            var utcDays = Math.floor(serial - 25569);
            var utcValue = utcDays * 86400;
            var dateInfo = new Date(utcValue * 1000);
            var year = dateInfo.getUTCFullYear();
            var month = dateInfo.getUTCMonth() + 1;
            var day = dateInfo.getUTCDate();
            return year + "-" + (month < 10 ? "0" + month : month) + "-" + (day < 10 ? "0" + day : day);
        }
    </script>
</body>

</html>