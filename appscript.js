// ============================================
// CONFIGURATION
// ============================================
const SHEET_NAMES = {
  responses: "Respuestas",
  analysis: "Análisis por Docente",
  comments: "Comentarios",
};

const QUESTIONS = [
  "¿Las clases del profesor son chéveres, divertidas, interesantes, y que facilitan mi aprendizaje?",
  "¿El docente permite realizar actividades de grupo en la clase?",
  "¿El docente exige el buen comportamiento para desarrollar la clase?",
  "¿La clase del docente se desarrolla con respeto y disciplina?",
  "¿El docente permite o me motiva permanente a que yo participe en clase?",
  "¿El profesor aclara dudas cuando lo solicitas?",
  "¿El docente es puntual en la llegada a clase?",
  "¿El docente atiende con respeto las inquietudes o dudas que se presentan?",
  "¿El docente responde tus preguntas con respeto y de manera oportuna?",
];

const ANSWER_VALUES = {
  Siempre: 4,
  "Casi siempre": 3,
  "Casi nunca": 2,
  Nunca: 1,
};

// ============================================
// WEB APP - SERVE DATA FOR DASHBOARD
// ============================================
function doGet(e) {
  // Check password
  const password = e.parameter.password;
  if (password !== "docentecolombo26") {
    return ContentService.createTextOutput(
      JSON.stringify({
        status: "error",
        message: "No autorizado",
      }),
    ).setMimeType(ContentService.MimeType.JSON);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const responsesSheet = ss.getSheetByName(SHEET_NAMES.responses);

  if (!responsesSheet || responsesSheet.getLastRow() <= 1) {
    return ContentService.createTextOutput(
      JSON.stringify({
        status: "success",
        data: [],
      }),
    ).setMimeType(ContentService.MimeType.JSON);
  }

  // Get data and convert to array of objects
  const data = responsesSheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  const formattedData = rows.map((row) => {
    const obj = {};
    headers.forEach((header, index) => {
      // Handle dates to string
      if (row[index] instanceof Date) {
        obj[header] = row[index].toISOString();
      } else {
        obj[header] = row[index];
      }
    });
    return obj;
  });

  return ContentService.createTextOutput(
    JSON.stringify({
      status: "success",
      data: formattedData,
    }),
  ).setMimeType(ContentService.MimeType.JSON);
}

// ============================================
// WEB APP - RECEIVE DATA FROM SURVEY
// ============================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // Save raw responses
    saveResponses(data);

    // Update analysis
    updateAnalysis();

    return ContentService.createTextOutput(
      JSON.stringify({
        status: "success",
        message: "Evaluación guardada exitosamente",
      }),
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    Logger.log("Error: " + error.toString());
    return ContentService.createTextOutput(
      JSON.stringify({
        status: "error",
        message: error.toString(),
      }),
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================
// SAVE RAW RESPONSES TO SHEET
// ============================================
function saveResponses(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAMES.responses);

  // Create sheet if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.responses);

    // Create headers
    const headers = [
      "Fecha y Hora",
      "Grupo",
      "Docente",
      "Materias",
      ...QUESTIONS,
      "Comentarios",
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet
      .getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#4285f4")
      .setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  }

  // Add each teacher's response
  data.responses.forEach((response) => {
    const row = [
      new Date(data.timestamp),
      data.group,
      response.teacher,
      response.subjects,
      ...response.answers,
      response.comments || "",
    ];
    sheet.appendRow(row);
  });

  // Auto-resize columns
  sheet.autoResizeColumns(1, sheet.getLastColumn());

  // Save comments separately
  saveComments(data);
}

// ============================================
// SAVE COMMENTS TO SEPARATE SHEET
// ============================================
function saveComments(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAMES.comments);

  // Create sheet if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.comments);

    const headers = [
      "Fecha y Hora",
      "Grupo",
      "Docente",
      "Materias",
      "Comentarios",
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet
      .getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#34a853")
      .setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  }

  // Add comments (only if not empty)
  data.responses.forEach((response) => {
    if (response.comments && response.comments.trim() !== "") {
      const row = [
        new Date(data.timestamp),
        data.group,
        response.teacher,
        response.subjects,
        response.comments,
      ];
      sheet.appendRow(row);
    }
  });

  sheet.autoResizeColumns(1, sheet.getLastColumn());
}

// ============================================
// UPDATE ANALYSIS SHEET WITH CHARTS
// ============================================
function updateAnalysis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const responsesSheet = ss.getSheetByName(SHEET_NAMES.responses);

  if (!responsesSheet || responsesSheet.getLastRow() <= 1) {
    return; // No data yet
  }

  let analysisSheet = ss.getSheetByName(SHEET_NAMES.analysis);

  // Create or clear analysis sheet
  if (!analysisSheet) {
    analysisSheet = ss.insertSheet(SHEET_NAMES.analysis);
  } else {
    analysisSheet.clear();
    // Remove all existing charts
    const charts = analysisSheet.getCharts();
    charts.forEach((chart) => analysisSheet.removeChart(chart));
  }

  // Get all responses
  const data = responsesSheet.getDataRange().getValues();
  const headers = data[0];
  const responses = data.slice(1);

  // Find column indices
  const teacherCol = headers.indexOf("Docente");
  const firstQuestionCol = 4; // Always starts at column E (index 4)

  // Group responses by teacher
  const teacherData = {};

  responses.forEach((row) => {
    const teacher = row[teacherCol];

    if (!teacherData[teacher]) {
      teacherData[teacher] = {
        responses: [],
        count: 0,
      };
    }

    const answers = row.slice(
      firstQuestionCol,
      firstQuestionCol + QUESTIONS.length,
    );
    teacherData[teacher].responses.push(answers);
    teacherData[teacher].count++;
  });

  // Create analysis for each teacher
  let currentRow = 1;

  Object.keys(teacherData)
    .sort()
    .forEach((teacher) => {
      const data = teacherData[teacher];

      // Calculate averages for each question
      const averages = QUESTIONS.map((_, qIndex) => {
        const sum = data.responses.reduce((acc, response) => {
          return acc + (ANSWER_VALUES[response[qIndex]] || 0);
        }, 0);
        return sum / data.count;
      });

      const overallAverage =
        averages.reduce((a, b) => a + b, 0) / averages.length;

      // Write teacher name and summary
      analysisSheet.getRange(currentRow, 1).setValue(`DOCENTE: ${teacher}`);
      analysisSheet
        .getRange(currentRow, 1)
        .setFontSize(14)
        .setFontWeight("bold")
        .setBackground("#fbbc04");

      analysisSheet
        .getRange(currentRow + 1, 1)
        .setValue("Número de evaluaciones:");
      analysisSheet.getRange(currentRow + 1, 2).setValue(data.count);

      analysisSheet.getRange(currentRow + 2, 1).setValue("Promedio General:");
      analysisSheet
        .getRange(currentRow + 2, 2)
        .setValue(overallAverage.toFixed(2));
      analysisSheet.getRange(currentRow + 2, 2).setNumberFormat("0.00");

      // Create data table for chart
      const chartStartRow = currentRow + 4;

      // Headers
      analysisSheet.getRange(chartStartRow, 1).setValue("Pregunta");
      analysisSheet.getRange(chartStartRow, 2).setValue("Promedio");
      analysisSheet
        .getRange(chartStartRow, 1, 1, 2)
        .setFontWeight("bold")
        .setBackground("#e8eaf6");

      // Data
      QUESTIONS.forEach((question, index) => {
        analysisSheet
          .getRange(chartStartRow + 1 + index, 1)
          .setValue(`P${index + 1}`);
        analysisSheet
          .getRange(chartStartRow + 1 + index, 2)
          .setValue(averages[index]);
      });

      // Format numbers
      analysisSheet
        .getRange(chartStartRow + 1, 2, QUESTIONS.length, 1)
        .setNumberFormat("0.00");

      // Create bar chart
      const chartRange = analysisSheet.getRange(
        chartStartRow,
        1,
        QUESTIONS.length + 1,
        2,
      );
      const chart = analysisSheet
        .newChart()
        .setChartType(Charts.ChartType.BAR)
        .addRange(chartRange)
        .setPosition(currentRow, 4, 0, 0)
        .setOption("title", `Evaluación: ${teacher}`)
        .setOption("width", 600)
        .setOption("height", 400)
        .setOption("hAxis", {
          title: "Promedio (1-4)",
          minValue: 0,
          maxValue: 4,
          gridlines: { count: 5 },
        })
        .setOption("vAxis", { title: "Preguntas" })
        .setOption("legend", { position: "none" })
        .setOption("colors", ["#4285f4"])
        .build();

      analysisSheet.insertChart(chart);

      // Create radar/spider chart for better visualization
      const radarChart = analysisSheet
        .newChart()
        .setChartType(Charts.ChartType.RADAR)
        .addRange(chartRange)
        .setPosition(currentRow, 14, 0, 0)
        .setOption("title", `Vista Radial: ${teacher}`)
        .setOption("width", 500)
        .setOption("height", 400)
        .setOption("legend", { position: "none" })
        .setOption("colors", ["#ea4335"])
        .build();

      analysisSheet.insertChart(radarChart);

      // CREATE PIE CHART - Distribution of responses per question
      // Calculate distribution of all answers for this teacher
      const distributionData = {
        Siempre: 0,
        "Casi siempre": 0,
        "Casi nunca": 0,
        Nunca: 0,
      };

      data.responses.forEach((response) => {
        response.forEach((answer) => {
          if (distributionData.hasOwnProperty(answer)) {
            distributionData[answer]++;
          }
        });
      });

      // Create pie chart data table
      const pieChartStartRow = currentRow;
      const pieChartStartCol = 24; // Column X

      analysisSheet
        .getRange(pieChartStartRow, pieChartStartCol)
        .setValue("Respuesta");
      analysisSheet
        .getRange(pieChartStartRow, pieChartStartCol + 1)
        .setValue("Cantidad");
      analysisSheet
        .getRange(pieChartStartRow, pieChartStartCol, 1, 2)
        .setFontWeight("bold")
        .setBackground("#e8eaf6");

      const pieData = [
        ["Siempre", distributionData["Siempre"]],
        ["Casi siempre", distributionData["Casi siempre"]],
        ["Casi nunca", distributionData["Casi nunca"]],
        ["Nunca", distributionData["Nunca"]],
      ];

      pieData.forEach((row, index) => {
        analysisSheet
          .getRange(pieChartStartRow + 1 + index, pieChartStartCol)
          .setValue(row[0]);
        analysisSheet
          .getRange(pieChartStartRow + 1 + index, pieChartStartCol + 1)
          .setValue(row[1]);
      });

      // Create pie chart
      const pieChartRange = analysisSheet.getRange(
        pieChartStartRow,
        pieChartStartCol,
        5,
        2,
      );
      const pieChart = analysisSheet
        .newChart()
        .setChartType(Charts.ChartType.PIE)
        .addRange(pieChartRange)
        .setPosition(currentRow + QUESTIONS.length + 2, 4, 0, 0)
        .setOption("title", `Distribución de Respuestas: ${teacher}`)
        .setOption("width", 500)
        .setOption("height", 400)
        .setOption("pieSliceText", "percentage")
        .setOption("colors", ["#34a853", "#fbbc04", "#ff6d01", "#ea4335"])
        .setOption("legend", { position: "right" })
        .setOption("pieHole", 0) // Set to 0.4 for donut chart
        .build();

      analysisSheet.insertChart(pieChart);

      // CREATE PERFORMANCE LEVEL PIE CHART
      // Categorize by performance level based on average
      const performanceLevels = {
        "Excelente (3.5-4.0)": 0,
        "Bueno (3.0-3.4)": 0,
        "Regular (2.5-2.9)": 0,
        "Bajo (<2.5)": 0,
      };

      averages.forEach((avg) => {
        if (avg >= 3.5) performanceLevels["Excelente (3.5-4.0)"]++;
        else if (avg >= 3.0) performanceLevels["Bueno (3.0-3.4)"]++;
        else if (avg >= 2.5) performanceLevels["Regular (2.5-2.9)"]++;
        else performanceLevels["Bajo (<2.5)"]++;
      });

      // Create performance level chart data
      const perfChartStartRow = currentRow;
      const perfChartStartCol = 30; // Column AD

      analysisSheet
        .getRange(perfChartStartRow, perfChartStartCol)
        .setValue("Nivel");
      analysisSheet
        .getRange(perfChartStartRow, perfChartStartCol + 1)
        .setValue("Preguntas");
      analysisSheet
        .getRange(perfChartStartRow, perfChartStartCol, 1, 2)
        .setFontWeight("bold")
        .setBackground("#e8eaf6");

      const perfData = [
        ["Excelente (3.5-4.0)", performanceLevels["Excelente (3.5-4.0)"]],
        ["Bueno (3.0-3.4)", performanceLevels["Bueno (3.0-3.4)"]],
        ["Regular (2.5-2.9)", performanceLevels["Regular (2.5-2.9)"]],
        ["Bajo (<2.5)", performanceLevels["Bajo (<2.5)"]],
      ];

      perfData.forEach((row, index) => {
        analysisSheet
          .getRange(perfChartStartRow + 1 + index, perfChartStartCol)
          .setValue(row[0]);
        analysisSheet
          .getRange(perfChartStartRow + 1 + index, perfChartStartCol + 1)
          .setValue(row[1]);
      });

      // Create performance level pie chart (donut style)
      const perfChartRange = analysisSheet.getRange(
        perfChartStartRow,
        perfChartStartCol,
        5,
        2,
      );
      const perfPieChart = analysisSheet
        .newChart()
        .setChartType(Charts.ChartType.PIE)
        .addRange(perfChartRange)
        .setPosition(currentRow + QUESTIONS.length + 2, 14, 0, 0)
        .setOption("title", `Nivel de Desempeño por Pregunta: ${teacher}`)
        .setOption("width", 500)
        .setOption("height", 400)
        .setOption("pieSliceText", "value")
        .setOption("colors", ["#34a853", "#fbbc04", "#ff6d01", "#ea4335"])
        .setOption("legend", { position: "bottom" })
        .setOption("pieHole", 0.4) // Donut chart
        .build();

      analysisSheet.insertChart(perfPieChart);

      // Move to next teacher section
      currentRow += QUESTIONS.length + 18; // Increased spacing for pie charts

      // Add separator
      analysisSheet.getRange(currentRow - 2, 1, 1, 20).setBackground("#f5f5f5");
    });

  // Auto-resize columns
  analysisSheet.autoResizeColumns(1, 2);

  // Add legend at the top
  analysisSheet.insertRowBefore(1);
  analysisSheet.getRange(1, 1).setValue("ANÁLISIS GLOBAL POR DOCENTE");
  analysisSheet
    .getRange(1, 1)
    .setFontSize(16)
    .setFontWeight("bold")
    .setBackground("#34a853")
    .setFontColor("#ffffff");
  analysisSheet.setColumnWidth(1, 400);
}

// ============================================
// MANUAL TRIGGER TO REFRESH ANALYSIS
// ============================================
function refreshAnalysis() {
  updateAnalysis();
  SpreadsheetApp.getUi().alert("Análisis actualizado exitosamente");
}

// ============================================
// CREATE MENU ON OPEN
// ============================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("📊 Evaluación Docente")
    .addItem("🔄 Actualizar Análisis", "refreshAnalysis")
    .addItem("📋 Ver Instrucciones", "showInstructions")
    .addToUi();
}

// ============================================
// SHOW INSTRUCTIONS
// ============================================
function showInstructions() {
  const ui = SpreadsheetApp.getUi();
  const instructions = `
INSTRUCCIONES DE USO:

1. Desplegar como Web App:
   - Ir a "Implementar" > "Nueva implementación"
   - Tipo: Aplicación web
   - Ejecutar como: Yo
   - Quién tiene acceso: Cualquier persona
   - Copiar la URL de la aplicación web

2. Pegar la URL en el código HTML:
   - Buscar: const SCRIPT_URL = 'YOUR_GOOGLE_APPS_SCRIPT_URL_HERE';
   - Reemplazar con tu URL

3. Las hojas se crearán automáticamente:
   - Respuestas: Datos brutos de cada evaluación
   - Análisis por Docente: Gráficos y promedios
   - Comentarios: Feedback escrito de estudiantes

4. Para actualizar los gráficos:
   - Menú: 📊 Evaluación Docente > 🔄 Actualizar Análisis
   - O esperar a que se actualice automáticamente con nuevas respuestas

ESCALA DE EVALUACIÓN:
Siempre = 4 puntos
Casi siempre = 3 puntos
Casi nunca = 2 puntos
Nunca = 1 punto

Promedio ideal: 3.5 - 4.0
Promedio bueno: 3.0 - 3.5
Requiere atención: < 3.0
  `;

  ui.alert("Instrucciones", instructions, ui.ButtonSet.OK);
}

// ============================================
// INITIAL SETUP FUNCTION (Run once)
// ============================================
function initialSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Create all sheets if they don't exist
  if (!ss.getSheetByName(SHEET_NAMES.responses)) {
    const sheet = ss.insertSheet(SHEET_NAMES.responses);
    const headers = [
      "Fecha y Hora",
      "Grupo",
      "Docente",
      "Materias",
      ...QUESTIONS,
      "Comentarios",
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet
      .getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#4285f4")
      .setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  }

  if (!ss.getSheetByName(SHEET_NAMES.comments)) {
    const sheet = ss.insertSheet(SHEET_NAMES.comments);
    const headers = [
      "Fecha y Hora",
      "Grupo",
      "Docente",
      "Materias",
      "Comentarios",
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet
      .getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#34a853")
      .setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  }

  if (!ss.getSheetByName(SHEET_NAMES.analysis)) {
    ss.insertSheet(SHEET_NAMES.analysis);
  }

  SpreadsheetApp.getUi().alert(
    "Setup completo",
    "Las hojas han sido creadas exitosamente.",
    SpreadsheetApp.getUi().ButtonSet.OK,
  );
}
