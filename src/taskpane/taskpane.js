Office.onReady(() => {
  // Constants
  const DEFAULT_RESULT_SHEET = "état complet";
  const TABLE_NAME_MAX_LENGTH = 255;
  const ROW_GAP = 3;
  const DEBOUNCE_DELAY = 300;


  function App() {
    const self = this;

    // Initialize all properties
    self.allSheets = [];
    self.sheets = [];
    self.listNominativeSheet = "";
    self.listNominativeTables = [];
    self.listNominativeTable = "";
    self.listNominativeColumns = [];
    self.listNominativeJoinColumn = "";
    self.selectedNominativeColumns = [];
    self.listEffetiveSheet = "";
    self.listEffetiveTables = [];
    self.listEffetiveTable = "";
    self.listEffetiveColumns = [];
    self.listEffetiveJoinColumn = "";
    self.selectedEffectiveColumns = [];
    self.showNominativeAdvanced = false;
    self.showEffectiveAdvanced = false;
    self.resultSheetName = DEFAULT_RESULT_SHEET;
    self.selectedNominativeGroupByColumns = [];
    self.listNominativeGroupByColumns = [];
    self.sheetCache = null;
    self.listenersAdded = false;
    self.filters = [];
    self.columnUniqueValues = new Map();
    self.showGroupBy = false;
    self.templateSheet = "";
    self.templateRangeAddress = "";
    self.showColumnOrdering = false;
    self.orderedColumns = [];
    self.extensionTables = [];
    self.showExtensionTables = false;

    // Logging utility
    self.log = (level, message, ...args) => {
      const timestamp = new Date().toISOString();
       console[level](`[${timestamp}] ${message}`, ...args);
    };

    // Debounce utility
    self.debounce = (func, wait) => {
      let timeout;
      return (...args) => {
        clearTimeout(timeout);
        timeout = setTimeout(() => func.apply(self, args), wait);
      };
    };

    // UI feedback utilities
    self.showError = (message) => {
      const root = document.getElementById("root");
      if (root) {
        const errorDiv = document.createElement("div");
        errorDiv.className = "notification is-danger";
        errorDiv.innerHTML = `<button class="delete" onclick="this.parentElement.remove()"></button>${message}`;
        root.insertBefore(errorDiv, root.firstChild);
      } else {
        alert(message);
      }
    };

    self.showProgress = (message) => {
      const root = document.getElementById("root");
      if (root) {
        const progressDiv = document.createElement("div");
        progressDiv.id = "progress";
        progressDiv.className = "notification is-info";
        progressDiv.textContent = message;
        root.insertBefore(progressDiv, root.firstChild);
      }
    };

    self.hideProgress = () => {
      const progress = document.getElementById("progress");
      if (progress) progress.remove();
    };

    // Input sanitization
    self.sanitizeSheetName = (name) => {
      return name.replace(/[:\\\/\*\?\[\]]/g, "_").slice(0, 31) || DEFAULT_RESULT_SHEET;
    };

    // Fetch unique values for a column (only from Nominative table for now)
    self.getUniqueColumnValues = async (columnName) => {
      if (self.columnUniqueValues.has(columnName)) {
        return self.columnUniqueValues.get(columnName);
      }
      try {
        const uniqueValues = await Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getItem(self.listNominativeSheet);
          const table = sheet.tables.getItem(self.listNominativeTable);
          const column = table.columns.getItem(columnName);
          const range = column.getDataBodyRange();
          range.load("values");
          await context.sync();

          const values = range.values.flat().map((val) => String(val || "").trim());
          const uniqueSet = new Set(values);
          return Array.from(uniqueSet).sort();
        });
        self.columnUniqueValues.set(columnName, uniqueValues);
        self.log("info", `Valeurs uniques récupérées pour la colonne "${columnName}":`, uniqueValues);
        return uniqueValues;
      } catch (error) {
        self.log("error", `Erreur lors de la récupération des valeurs uniques pour "${columnName}":`, error);
        self.showError(`Impossible de charger les valeurs pour la colonne "${columnName}".`);
        return [];
      }
    };

    // Enhanced header generation function
    self.generateTableHeader = (templateValues, tableInfo, groupRows, tableHeaders) => {
      const computedValues = {
        tableName: tableInfo.finalTableName,
        rowCount: tableInfo.rowCount,
        uniqueJoinValues: tableInfo.uniqueJoinValues || "N/A"
      };

      const sumColumns = new Set();
      const avgColumns = new Set();
      const expressionRegex = /\{([^}]+)\}/g;

      templateValues.forEach((row) => {
        row.forEach((cell) => {
          let match;
          while ((match = expressionRegex.exec(cell)) !== null) {
            const expression = match[1];
            if (expression.startsWith("sum:")) {
              sumColumns.add(expression.split(":")[1]);
            } else if (expression.startsWith("avg:")) {
              avgColumns.add(expression.split(":")[1]);
            }
          }
        });
      });

      const columnIndices = {};
      tableHeaders.forEach((header, index) => (columnIndices[header] = index));

      const tableData = groupRows.map((row) => tableHeaders.map((header) => row[header] || ""));
      sumColumns.forEach((colName) => {
        const colIndex = columnIndices[colName];
        if (colIndex !== undefined) {
          const values = tableData.map((row) => Number(row[colIndex]) || 0);
          computedValues[`sum:${colName}`] = values.reduce((a, b) => a + b, 0);
        } else {
          computedValues[`sum:${colName}`] = `N/A (Colonne "${colName}" non trouvée dans la sortie)`;
        }
      });
      avgColumns.forEach((colName) => {
        const colIndex = columnIndices[colName];
        if (colIndex !== undefined) {
          const values = tableData.map((row) => Number(row[colIndex]) || 0);
          computedValues[`avg:${colName}`] = values.length > 0 ? values.reduce((a, b) => a + b, 0) / values.length : 0;
        } else {
          computedValues[`avg:${colName}`] = `N/A (Colonne "${colName}" non trouvée dans la sortie)`;
        }
      });

      const processedValues = templateValues.map((row) =>
        row.map((cell) => {
          if (typeof cell !== "string") return cell;
          return cell.replace(expressionRegex, (match, expression) => {
            const parts = expression.split(":");
            if (parts.length === 2 && !isNaN(parts[1])) {
              const columnName = parts[0];
              const index = parseInt(parts[1], 10);
              if (index < 0 || index >= groupRows.length) return `N/A (Index ${index} hors plage)`;
              const rowData = groupRows[index];
              return rowData[columnName] !== undefined
                ? rowData[columnName]
                : `N/A (Colonne "${columnName}" non trouvée)`;
            }
            return computedValues[expression] ?? match;
          });
        })
      );

      return processedValues;
    };

    // State validation
    self.validateState = (nominativeHeaders, effectiveHeaders, allGroupByColumns) => {
      const errors = [];
      if (!self.listNominativeSheet || !self.listNominativeTable)
        errors.push("Feuille ou tableau nominatif non sélectionné.");
      if (!self.listEffetiveSheet || !self.listEffetiveTable)
        errors.push("Feuille ou tableau effectif non sélectionné.");
      if (nominativeHeaders.indexOf(self.listNominativeJoinColumn) === -1)
        errors.push(`Colonne de jointure nominative "${self.listNominativeJoinColumn}" non trouvée.`);
      if (effectiveHeaders.indexOf(self.listEffetiveJoinColumn) === -1)
        errors.push(`Colonne de jointure effective "${self.listEffetiveJoinColumn}" non trouvée.`);
      if (self.selectedNominativeColumns.length === 0) errors.push("Aucune colonne nominative sélectionnée.");
      if (self.selectedNominativeGroupByColumns.some((col) => !allGroupByColumns.includes(col)))
        errors.push("Colonnes de regroupement invalides sélectionnées.");
      self.filters.forEach((filter, index) => {
        if (filter.value.trim() && nominativeHeaders.indexOf(filter.column) === -1) {
          errors.push(`Filtre ${index + 1} colonne "${filter.column}" non trouvée.`);
        }
      });
      self.extensionTables.forEach((ext, index) => {
        if (!ext.sheet || !ext.table)
          errors.push(`Tableau d'extension ${index + 1} : Feuille ou tableau non sélectionné.`);
        if (!ext.joinColumnResult || self.orderedColumns.indexOf(ext.joinColumnResult) === -1)
          errors.push(
            `Tableau d'extension ${index + 1} : Colonne de jointure résultat "${ext.joinColumnResult}" non trouvée.`
          );
        if (!ext.joinColumnExtension || ext.columns.every((col) => col.name !== ext.joinColumnExtension))
          errors.push(
            `Tableau d'extension ${index + 1} : Colonne de jointure extension "${ext.joinColumnExtension}" non trouvée.`
          );
      });
      if (errors.length > 0) throw new Error(errors.join(" "));
      if (self.selectedEffectiveColumns.length === 0)
        self.log("warn", "Aucune colonne effective sélectionnée ; traitement avec les données nominatives uniquement");
    };

    self.process = async () => {
      const processButton = document.querySelector(".button.is-primary");
      if (!processButton) return;
    
      processButton.disabled = true;
      processButton.classList.add("is-loading");
      processButton.innerHTML = `
        <span class="icon">
          <i class="fas fa-spinner fa-spin"></i>
        </span>
        <span>Traitement en cours...</span>
      `;
    
      try {
        self.showProgress("Traitement des données...");
        const resultSheetInput = document.getElementById("resultSheetName");
        self.resultSheetName = self.sanitizeSheetName(
          resultSheetInput ? resultSheetInput.value.trim() : DEFAULT_RESULT_SHEET
        );
    
        await Excel.run(async (context) => {
          const nominativeSheet = context.workbook.worksheets.getItem(self.listNominativeSheet);
          const nominativeTable = nominativeSheet.tables.getItem(self.listNominativeTable);
          const effectiveSheet = context.workbook.worksheets.getItem(self.listEffetiveSheet);
          const effectiveTable = effectiveSheet.tables.getItem(self.listEffetiveTable);
          const resultSheet = context.workbook.worksheets.getItemOrNullObject(self.resultSheetName);
          const allTables = context.workbook.tables;
    
          const nominativeRange = nominativeTable.getRange();
          const effectiveRange = effectiveTable.getRange();
          nominativeRange.load("values");
          effectiveRange.load("values");
          allTables.load("items/name");
    
          const nominativeColumns = nominativeTable.columns;
          const effectiveColumns = effectiveTable.columns;
          nominativeColumns.load("items/name, items/numberFormat");
          effectiveColumns.load("items/name, items/numberFormat");
    
          const extensionData = {};
          const extensionFormats = {};
          for (const ext of self.extensionTables) {
            const sheet = context.workbook.worksheets.getItem(ext.sheet);
            const table = sheet.tables.getItem(ext.table);
            const range = table.getRange();
            range.load("values");
            const columns = table.columns;
            columns.load("items/name, items/numberFormat");
            extensionData[ext.table] = { range, columns };
          }
    
          let templateRange;
          if (!self.templateSheet || !self.templateRangeAddress) {
            try {
              templateRange = context.workbook.names.getItem("HeaderTemplate").getRange();
            } catch (e) {
              throw new Error(
                "Veuillez spécifier la feuille et la plage du modèle, ou définir une plage nommée 'HeaderTemplate'."
              );
            }
          } else {
            const templateSheet = context.workbook.worksheets.getItem(self.templateSheet);
            templateRange = templateSheet.getRange(self.templateRangeAddress);
          }
          templateRange.load("values, rowCount, columnCount");
    
          await context.sync();
    
          const templateValues = templateRange.values;
          const templateRowCount = templateRange.rowCount;
          const templateColumnCount = templateRange.columnCount;
    
          const nominativeData = nominativeRange.values;
          const nominativeHeaders = nominativeData[0];
          const nominativeRows = nominativeData.slice(1);
          const effectiveData = effectiveRange.values;
          const effectiveHeaders = effectiveData[0];
          const effectiveRows = effectiveData.slice(1);
    
          for (const [tableName, { range, columns }] of Object.entries(extensionData)) {
            const extData = range.values;
            extensionData[tableName].headers = extData[0];
            extensionData[tableName].rows = extData.slice(1);
            extensionFormats[tableName] = {};
            columns.items.forEach((col) => {
              const format =
                col.numberFormat && col.numberFormat[0] && col.numberFormat[0][0] ? col.numberFormat[0][0] : "General";
              extensionFormats[tableName][col.name] = format;
            });
          }
    
          const allGroupByColumns = [
            ...nominativeHeaders,
            ...effectiveHeaders,
            ...self.extensionTables.flatMap((ext) => ext.columns.map((col) => col.name))
          ];
          self.validateState(nominativeHeaders, effectiveHeaders, allGroupByColumns);
    
          const nominativeFormats = {};
          nominativeColumns.items.forEach((col) => {
            nominativeFormats[col.name] = col.numberFormat?.[0]?.[0] || "General";
          });
          const effectiveFormats = {};
          effectiveColumns.items.forEach((col) => {
            effectiveFormats[col.name] = col.numberFormat?.[0]?.[0] || "General";
          });
    
          const effectiveLookup = new Map();
          effectiveRows.forEach((row) => {
            const joinValue = String(row[effectiveHeaders.indexOf(self.listEffetiveJoinColumn)] || "").trim();
            effectiveLookup.set(joinValue, row);
          });
    
          const extensionLookups = {};
          self.extensionTables.forEach((ext) => {
            const extRows = extensionData[ext.table].rows;
            const extHeaders = extensionData[ext.table].headers;
            const joinColIndex = extHeaders.indexOf(ext.joinColumnExtension);
            const lookup = new Map();
            extRows.forEach((row) => {
              const joinValue = String(row[joinColIndex] || "").trim();
              lookup.set(joinValue, row);
            });
            extensionLookups[ext.table] = { lookup, headers: extHeaders };
          });
    
          let filteredNominativeRows = nominativeRows;
          const activeFilters = self.filters.filter((f) => f.value.trim() !== "");
          if (activeFilters.length > 0) {
            const filterColIndices = activeFilters.map((f) => nominativeHeaders.indexOf(f.column));
            filteredNominativeRows = nominativeRows.filter((row) =>
              activeFilters.every(
                (f, i) =>
                  String(row[filterColIndices[i]] || "")
                    .trim()
                    .toLowerCase() === f.value.toLowerCase()
              )
            );
          }
    
          const combinedData = filteredNominativeRows.map((row) => {
            const nominativeJoinValue = String(
              row[nominativeHeaders.indexOf(self.listNominativeJoinColumn)] || ""
            ).trim();
            const effectiveRow = effectiveLookup.get(nominativeJoinValue) || [];
    
            const valueMap = {};
            nominativeHeaders.forEach((col, idx) => {
              valueMap[col] = row[idx] ?? "";
            });
            effectiveHeaders.forEach((col, idx) => {
              valueMap[col] = effectiveRow[idx] ?? "";
            });
            self.extensionTables.forEach((ext) => {
              const extLookup = extensionLookups[ext.table].lookup;
              const extHeaders = extensionLookups[ext.table].headers;
              const resultJoinValue = valueMap[ext.joinColumnResult] ?? "";
              const extRow = extLookup.get(String(resultJoinValue).trim()) || [];
              extHeaders.forEach((col, idx) => {
                valueMap[col] = extRow[idx] ?? "";
              });
            });
    
            return valueMap;
          });
    
          const groupedData = {};
          combinedData.forEach((row) => {
            const groupKey = self.selectedNominativeGroupByColumns
              .map((col) => String(row[col] || "").trim())
              .join("|");
            if (!groupedData[groupKey]) groupedData[groupKey] = [];
            groupedData[groupKey].push(row);
          });
    
          // Fix: Use a separate variable for the result sheet
          let targetSheet;
          if (resultSheet.isNullObject) {
            targetSheet = context.workbook.worksheets.add(self.resultSheetName);
          } else {
            targetSheet = resultSheet;
            targetSheet.getUsedRange()?.clear();
          }
          targetSheet.activate();
    
          let currentRow = 1;
          let tableCounter = 1;
          const existingTableNames = new Set(allTables.items.map((t) => t.name.toLowerCase()));
          const operations = [];
    
          for (const [groupKey, groupRows] of Object.entries(groupedData)) {
            const tableHeaders = self.orderedColumns;
            const tableData = groupRows.map((row) => tableHeaders.map((header) => row[header] || ""));
    
            const rowCount = tableData.length;
            let finalTableName = `Tableau_${tableCounter}`;
            while (existingTableNames.has(finalTableName.toLowerCase())) {
              tableCounter++;
              finalTableName = `Tableau_${tableCounter}`;
            }
    
            operations.push({
              currentRow,
              finalTableName,
              rowCount,
              uniqueJoinValues: new Set(groupRows.map((r) => String(r[self.listNominativeJoinColumn] || "").trim()))
                .size,
              tableHeaders,
              tableData,
              groupRows,
              execute: function() {
                const processedHeaderValues = self.generateTableHeader(
                  templateValues,
                  this,
                  this.groupRows,
                  this.tableHeaders
                );
                const headerRange = targetSheet.getRangeByIndexes(
                  this.currentRow - 1,
                  0,
                  templateRowCount,
                  templateColumnCount
                );
                headerRange.copyFrom(templateRange, Excel.RangeCopyType.formats);
                headerRange.values = processedHeaderValues;
    
                const tableStartRow = this.currentRow + templateRowCount;
                const tableRange = targetSheet.getRangeByIndexes(
                  tableStartRow - 1,
                  0,
                  this.tableData.length + 1,
                  this.tableHeaders.length
                );
                tableRange.values = [this.tableHeaders, ...this.tableData];
    
                const headerTableRange = targetSheet.getRangeByIndexes(
                  tableStartRow - 1,
                  0,
                  1,
                  this.tableHeaders.length
                );
                const formats = this.tableHeaders.map(
                  (header) =>
                    nominativeFormats[header] ||
                    effectiveFormats[header] ||
                    self.extensionTables.reduce((fmt, ext) => fmt || extensionFormats[ext.table][header], null) ||
                    "General"
                );
                headerTableRange.numberFormat = [formats];
    
                const newTable = targetSheet.tables.add(tableRange, true);
                newTable.name = this.finalTableName;
                existingTableNames.add(this.finalTableName.toLowerCase());
                return tableStartRow + this.tableData.length + ROW_GAP;
              }
            });
    
            currentRow += operations[operations.length - 1].rowCount + ROW_GAP + templateRowCount;
            tableCounter++;
          }
    
          let nextRow = 1;
          for (const op of operations) {
            op.currentRow = nextRow;
            nextRow = op.execute();
          }
    
          targetSheet.getUsedRange().format.autofitColumns();
          await context.sync();
          self.log("info", `Feuille '${self.resultSheetName}' créée avec ${tableCounter - 1} tableaux`);
          self.hideProgress();
        });
      } catch (error) {
        self.hideProgress();
        self.log("error", "Erreur de traitement:", error);
        self.showError(`Une erreur est survenue : ${error.message}`);
      } finally {
        processButton.disabled = false;
        processButton.classList.remove("is-loading");
        processButton.innerHTML = "Traiter les Données";
      }
    };

    // Update ordered columns to include only selected columns for output
    self.updateOrderedColumns = () => {
      const combined = [
        ...self.selectedNominativeColumns,
        ...self.selectedEffectiveColumns,
        ...self.extensionTables.flatMap((ext) => ext.selectedColumns)
      ];
      const currentSet = new Set(combined);
      const filteredOrdered = self.orderedColumns.filter((col) => currentSet.has(col));
      const newColumns = combined.filter((col) => !self.orderedColumns.includes(col));
      self.orderedColumns = [...filteredOrdered, ...newColumns];

      if (self.showColumnOrdering) {
        self.renderColumnOrderList();
      }
    };

    // Update group-by columns to include all columns from source tables
    self.updateGroupByColumnsList = () => {
      const allColumns = [
        ...(self.listNominativeColumns || []).map((col) => col.name),
        ...(self.listEffetiveColumns || []).map((col) => col.name),
        ...self.extensionTables.flatMap((ext) => (ext.columns || []).map((col) => col.name))
      ];
      self.listNominativeGroupByColumns = Array.from(new Set(allColumns)).map((name) => ({ name }));
      self.selectedNominativeGroupByColumns = self.selectedNominativeGroupByColumns.filter((col) =>
        self.listNominativeGroupByColumns.some((c) => c.name === col)
      );
      if (
        self.selectedNominativeGroupByColumns.length === 0 &&
        self.listNominativeGroupByColumns.length > 0 &&
        self.listNominativeJoinColumn
      ) {
        self.selectedNominativeGroupByColumns = [self.listNominativeJoinColumn];
      }
    };

    self.renderColumnOrderList = () => {
      const columnOrderList = document.getElementById("columnOrderList");
      if (columnOrderList) {
        columnOrderList.innerHTML = self.orderedColumns
          .map(
            (col) => `
            <div class="column-order-item" draggable="true" data-column="${col}">
              <span class="drag-handle">☰</span> ${col}
            </div>
          `
          )
          .join("");
        self.setupDragAndDrop();
      }
    };

    self.setupDragAndDrop = () => {
      const columnOrderList = document.getElementById("columnOrderList");
      if (!columnOrderList) return;

      let draggedItem = null;

      const getDragAfterElement = (container, y) => {
        const draggableElements = [...container.querySelectorAll(".column-order-item:not(.dragging)")];
        return draggableElements.reduce(
          (closest, child) => {
            const box = child.getBoundingClientRect();
            const offset = y - box.top - box.height / 2;
            return offset < 0 && offset > closest.offset ? { offset: offset, element: child } : closest;
          },
          { offset: Number.NEGATIVE_INFINITY }
        ).element;
      };

      columnOrderList.querySelectorAll(".column-order-item").forEach((item) => {
        item.addEventListener("dragstart", (e) => {
          draggedItem = e.target;
          setTimeout(() => item.classList.add("dragging"), 0);
        });

        item.addEventListener("dragend", () => {
          item.classList.remove("dragging");
        });

        item.addEventListener("dragover", (e) => {
          e.preventDefault();
          const afterElement = getDragAfterElement(columnOrderList, e.clientY);
          if (afterElement == null) {
            columnOrderList.appendChild(draggedItem);
          } else {
            columnOrderList.insertBefore(draggedItem, afterElement);
          }
        });
      });

      columnOrderList.addEventListener("dragend", () => {
        const newOrder = Array.from(columnOrderList.children).map((item) => item.dataset.column);
        self.orderedColumns = newOrder;
        self.log("info", "Colonnes réordonnées:", self.orderedColumns);
      });
    };

    self.toggleColumnOrdering = () => {
      self.showColumnOrdering = !self.showColumnOrdering;
      self.render();
    };

    self.resetColumnOrder = () => {
      self.orderedColumns = [
        ...self.selectedNominativeColumns,
        ...self.selectedEffectiveColumns,
        ...self.extensionTables.flatMap((ext) => ext.selectedColumns)
      ];
      if (self.showColumnOrdering) {
        self.renderColumnOrderList();
      }
    };

    self.getWorkbookSheets = async () => {
      if (self.sheetCache) return self.sheetCache;
      try {
        const sheets = await Excel.run(async (context) => {
          const sheets = context.workbook.worksheets;
          sheets.load("items/name");
          await context.sync();

          for (let i = 0; i < sheets.items.length; i++) {
            const sheet = sheets.items[i];
            const tables = sheet.tables;
            tables.load("items/name");
            await context.sync();

            for (let j = 0; j < tables.items.length; j++) {
              const table = tables.items[j];
              table.columns.load("items/name");
              
            }
            await context.sync();
          }
          return sheets;
        });
        self.sheetCache = sheets;
        return sheets;
      } catch (error) {
        self.log("error", "Erreur lors de la récupération des feuilles du classeur:", error);
        self.showError("Impossible de charger les feuilles du classeur.");
        return { items: [] };
      }
    };

    self.invalidateCache = () => {
      self.sheetCache = null;
      self.columnUniqueValues.clear();
    };

    self.updateTemplateRange = async (sheetName) => {
      try {
        await Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getItem(sheetName);
          const usedRange = sheet.getUsedRange(true);
          usedRange.load("address");
          await context.sync();

          if (usedRange.address) {
            self.templateRangeAddress = usedRange.address.split("!")[1] || "A1";
          } else {
            self.templateRangeAddress = "A1";
          }

          const rangeInput = document.getElementById("templateRangeAddress");
          if (rangeInput) rangeInput.value = self.templateRangeAddress;
        });
      } catch (error) {
        self.templateRangeAddress = "A1";
        self.showError(`Impossible de détecter la plage utilisée pour "${sheetName}". Par défaut à A1.`);
        const rangeInput = document.getElementById("templateRangeAddress");
        if (rangeInput) rangeInput.value = self.templateRangeAddress;
      }
    };

    self.init = async () => {
      try {
        self.filters = [];
        self.extensionTables = [];


        console.log("befor get sheet name");
        
        const workbookSheets = await self.getWorkbookSheets();

        console.log("after get sheet name");
        self.allSheets = workbookSheets.items;
        self.sheets = workbookSheets.items.filter((sheet) => sheet.tables.items.length > 0);

        if (self.sheets.length === 0) {
          self.showError(
            "Aucune feuille avec des tableaux trouvée dans le classeur pour les listes Nominative ou Effective."
          );
          return;
        }
        if (self.allSheets.length === 0) {
          self.showError("Aucune feuille trouvée dans le classeur.");
          return;
        }

        self.listNominativeSheet = self.sheets[0].name;
        self.listEffetiveSheet = self.sheets[0].name;
        self.templateSheet = self.allSheets[0].name;
        await self.updateTemplateRange(self.templateSheet);
        self.handleNominativeSheetChange();
        self.handleEffectiveSheetChange();
      } catch (error) {
        console.log();
        
        self.log("error", "Erreur d'initialisation:", error);
        self.showError(`Échec de l'initialisation : ${error.message}`);
      }
    };

    self.handleNominativeSheetChange = () => {
      const nominativeSheet = self.sheets.find((sheet) => sheet.name === self.listNominativeSheet);
      if (nominativeSheet) {
        self.listNominativeTables = nominativeSheet.tables.items;
        if (self.listNominativeTables.length > 0) {
          self.listNominativeTable = self.listNominativeTables[0].name;
          self.handleNominativeTableChange();
        } else {
          self.listNominativeTables = [];
          self.listNominativeTable = "";
          self.listNominativeColumns = [];
          self.selectedNominativeColumns = [];
          self.filters = [];
          self.columnUniqueValues.clear();
          self.updateGroupByColumnsList();
          self.orderedColumns = [];
        }
        self.updateDropdowns();
      }
    };

    self.handleNominativeTableChange = () => {
      const nominativeTable = self.listNominativeTables.find((table) => table.name === self.listNominativeTable);
      if (nominativeTable) {
        self.listNominativeColumns = nominativeTable.columns.items;
        if (self.listNominativeColumns.length > 0) {
          self.listNominativeJoinColumn = self.listNominativeColumns[0].name;
          self.selectedNominativeColumns = self.listNominativeColumns.map((column) => column.name);
          self.filters = [];
          self.columnUniqueValues.clear();
          self.updateGroupByColumnsList();
          self.orderedColumns = [...self.selectedNominativeColumns];
        } else {
          self.listNominativeJoinColumn = "";
          self.selectedNominativeColumns = [];
          self.filters = [];
          self.columnUniqueValues.clear();
          self.updateGroupByColumnsList();
          self.orderedColumns = [];
        }
        self.updateDropdowns();
        self.renderGroupBySummary();
      }
    };

    self.handleNominativeJoinColumnChange = () => {
      if (self.listNominativeGroupByColumns.some((col) => col.name === self.listNominativeJoinColumn)) {
        self.selectedNominativeGroupByColumns = [self.listNominativeJoinColumn];
        self.updateGroupByColumns(
          "nominativeGroupByColumns",
          self.listNominativeGroupByColumns,
          self.selectedNominativeGroupByColumns
        );
        self.renderGroupBySummary();
      }
    };

    self.handleEffectiveSheetChange = () => {
      const effectiveSheet = self.sheets.find((sheet) => sheet.name === self.listEffetiveSheet);
      if (effectiveSheet) {
        self.listEffetiveTables = effectiveSheet.tables.items;
        if (self.listEffetiveTables.length > 0) {
          self.listEffetiveTable = self.listEffetiveTables[0].name;
          self.handleEffectiveTableChange();
        } else {
          self.listEffetiveTables = [];
          self.listEffetiveTable = "";
          self.listEffetiveColumns = [];
          self.selectedEffectiveColumns = [];
          self.updateOrderedColumns();
          self.updateGroupByColumnsList();
        }
        self.updateDropdowns();
      }
    };

    self.handleEffectiveTableChange = () => {
      const effectiveTable = self.listEffetiveTables.find((table) => table.name === self.listEffetiveTable);
      if (effectiveTable) {
        self.listEffetiveColumns = effectiveTable.columns.items;
        if (self.listEffetiveColumns.length > 0) {
          self.listEffetiveJoinColumn = self.listEffetiveColumns[0].name;
          self.selectedEffectiveColumns = self.listEffetiveColumns.map((column) => column.name);
          self.updateOrderedColumns();
          self.updateGroupByColumnsList();
        } else {
          self.listEffetiveJoinColumn = "";
          self.selectedEffectiveColumns = [];
          self.updateOrderedColumns();
          self.updateGroupByColumnsList();
        }
        self.updateDropdowns();
      }
    };

    self.handleEffectiveJoinColumnChange = () => {
      self.updateDropdowns();
    };

    self.handleExtensionSheetChange = (index) => {
      const ext = self.extensionTables[index];
      const sheet = self.sheets.find((s) => s.name === ext.sheet);
      if (sheet) {
        ext.tables = sheet.tables.items;
        ext.table = ext.tables.length > 0 ? ext.tables[0].name : "";
        self.handleExtensionTableChange(index);
      }
      self.render();
    };

    self.handleExtensionTableChange = (index) => {
      const ext = self.extensionTables[index];
      const table = ext.tables.find((t) => t.name === ext.table);
      if (table) {
        ext.columns = table.columns.items;
        ext.joinColumnExtension = ext.columns.length > 0 ? ext.columns[0].name : "";
        ext.selectedColumns = ext.columns.map((col) => col.name);
        ext.joinColumnResult = self.orderedColumns.length > 0 ? self.orderedColumns[0] : "";
        ext.showColumns = false;
      } else {
        ext.columns = [];
        ext.joinColumnExtension = "";
        ext.selectedColumns = [];
        ext.joinColumnResult = "";
        ext.showColumns = false;
      }
      self.updateOrderedColumns();
      self.updateGroupByColumnsList();
      self.render();
    };

    self.addExtensionTable = () => {
      if (self.sheets.length > 0) {
        const newExt = {
          sheet: self.sheets[0].name,
          table: "",
          joinColumnResult: self.orderedColumns.length > 0 ? self.orderedColumns[0] : "",
          joinColumnExtension: "",
          selectedColumns: [],
          tables: [],
          columns: [],
          showColumns: false
        };
        self.extensionTables.push(newExt);
        self.handleExtensionSheetChange(self.extensionTables.length - 1);
      }
    };

    self.removeExtensionTable = (index) => {
      self.extensionTables.splice(index, 1);
      self.updateOrderedColumns();
      self.updateGroupByColumnsList();
      self.render();
    };

    self.toggleExtensionColumns = (index) => {
      self.extensionTables[index].showColumns = !self.extensionTables[index].showColumns;
      self.render();
    };

    self.updateDropdowns = () => {
      self.updateDropdown("listNominativeSheet", self.sheets, self.listNominativeSheet);
      self.updateDropdown("listNominativeTable", self.listNominativeTables, self.listNominativeTable);
      self.updateDropdown("listNominativeJoinColumn", self.listNominativeColumns, self.listNominativeJoinColumn);
      self.updateDropdown("listEffetiveSheet", self.sheets, self.listEffetiveSheet);
      self.updateDropdown("listEffetiveTable", self.listEffetiveTables, self.listEffetiveTable);
      self.updateDropdown("listEffetiveJoinColumn", self.listEffetiveColumns, self.listEffetiveJoinColumn);
      self.updateDropdown("templateSheet", self.allSheets, self.templateSheet);
      self.updateColumnSelection("nominativeColumns", self.listNominativeColumns, self.selectedNominativeColumns);
      self.updateColumnSelection("effectiveColumns", self.listEffetiveColumns, self.selectedEffectiveColumns);
      self.updateGroupByColumns(
        "nominativeGroupByColumns",
        self.listNominativeGroupByColumns,
        self.selectedNominativeGroupByColumns
      );
      self.renderGroupBySummary();
    };

    self.updateDropdown = (id, items, selectedValue) => {
      const dropdown = document.getElementById(id);
      if (dropdown) {
        dropdown.innerHTML = items
          .map(
            (item) =>
              `<option value="${item.name}" ${item.name === selectedValue ? "selected" : ""}>${item.name}</option>`
          )
          .join("");
      }
    };

    self.updateColumnSelection = (id, columns, selectedColumns) => {
      const container = document.getElementById(id);
      if (container) {
        container.innerHTML = columns
          .map(
            (column) => `
            <div class="field">
              <div class="control">
                <label class="checkbox">
                  <input type="checkbox" value="${column.name}" ${
              selectedColumns.includes(column.name) ? "checked" : ""
            }>
                  ${column.name}
                </label>
              </div>
            </div>
          `
          )
          .join("");

        container.querySelectorAll('input[type="checkbox"]').forEach((checkbox) => {
          checkbox.addEventListener(
            "change",
            self.debounce((e) => {
              const columnName = e.target.value;
              if (e.target.checked) {
                if (!selectedColumns.includes(columnName)) selectedColumns.push(columnName);
              } else {
                const index = selectedColumns.indexOf(columnName);
                if (index !== -1) selectedColumns.splice(index, 1);
              }
              self.updateOrderedColumns();
              self.updateGroupByColumnsList();
              if (self.showColumnOrdering) {
                self.renderColumnOrderList();
              }
            }, DEBOUNCE_DELAY)
          );
        });

        const selectAllBtn = document.getElementById(`${id}-select-all`);
        const unselectAllBtn = document.getElementById(`${id}-unselect-all`);
        if (selectAllBtn) {
          selectAllBtn.addEventListener("click", () => {
            selectedColumns.length = 0;
            selectedColumns.push(...columns.map((col) => col.name));
            self.updateOrderedColumns();
            self.updateGroupByColumnsList();
            if (self.showColumnOrdering) {
              self.renderColumnOrderList();
            }
            self.updateColumnSelection(id, columns, selectedColumns);
          });
        }
        if (unselectAllBtn) {
          unselectAllBtn.addEventListener("click", () => {
            selectedColumns.length = 0;
            self.updateOrderedColumns();
            self.updateGroupByColumnsList();
            if (self.showColumnOrdering) {
              self.renderColumnOrderList();
            }
            self.updateColumnSelection(id, columns, selectedColumns);
          });
        }
      }
    };

    self.updateGroupByColumns = (id, columns, selectedColumns) => {
      const container = document.getElementById(id);
      if (container) {
        container.innerHTML = columns
          .map(
            (column) => `
            <div class="field">
              <div class="control">
                <label class="checkbox">
                  <input type="checkbox" value="${column.name}" ${
              selectedColumns.includes(column.name) ? "checked" : ""
            }>
                  ${column.name}
                </label>
              </div>
            </div>
          `
          )
          .join("");

        container.querySelectorAll('input[type="checkbox"]').forEach((checkbox) => {
          checkbox.addEventListener(
            "change",
            self.debounce((e) => {
              const columnName = e.target.value;
              if (e.target.checked) {
                if (!selectedColumns.includes(columnName)) selectedColumns.push(columnName);
              } else {
                const index = selectedColumns.indexOf(columnName);
                if (index !== -1) selectedColumns.splice(index, 1);
              }
              self.renderGroupBySummary();
            }, DEBOUNCE_DELAY)
          );
        });
      }
    };

    self.renderGroupBySummary = () => {
      const groupByHeader = document.querySelector(".accordion-header .group-by-summary");
      if (groupByHeader) {
        const summary =
          self.selectedNominativeGroupByColumns.length > 0
            ? ` (${self.selectedNominativeGroupByColumns.join(", ")})`
            : " (Aucun)";
        groupByHeader.textContent = `Colonnes de Regroupement${summary}`;
      }
    };

    self.toggleAdvanced = (section) => {
      if (section === "nominative") {
        self.showNominativeAdvanced = !self.showNominativeAdvanced;
      } else if (section === "effective") {
        self.showEffectiveAdvanced = !self.showEffectiveAdvanced;
      }
      self.render();
    };

    self.toggleGroupBy = () => {
      self.showGroupBy = !self.showGroupBy;
      self.render();
    };

    self.toggleExtensionTables = () => {
      self.showExtensionTables = !self.showExtensionTables;
      self.render();
    };

    self.render = async () => {
      try {
        const root = document.getElementById("root");
        if (!root) throw new Error("Élément racine non trouvé dans le DOM");

        for (const filter of self.filters) {
          if (!self.columnUniqueValues.has(filter.column)) {
            await self.getUniqueColumnValues(filter.column);
          }
        }

        const groupBySummary =
          self.selectedNominativeGroupByColumns.length > 0
            ? ` (${self.selectedNominativeGroupByColumns.join(", ")})`
            : " (Aucun)";

        root.innerHTML = `
          <div class="m-3 box p-4">
            <h1 class="title is-4">Outil de Liaison de Tableaux Excel</h1>
            <div class="columns is-mobile">
              <div class="column">
                <h2 class="subtitle is-5 has-text-weight-bold">Liste Nominative</h2>
                <div class="field">
                  <label class="label" for="listNominativeSheet" data-tooltip="Sélectionnez la feuille contenant votre tableau de données principal">Feuille :</label>
                  <div class="control">
                    <div class="select is-fullwidth">
                      <select name="listNominativeSheet" id="listNominativeSheet">
                        ${self.sheets
                          .map(
                            (sheet) =>
                              `<option value="${sheet.name}" ${
                                sheet.name === self.listNominativeSheet ? "selected" : ""
                              }>${sheet.name}</option>`
                          )
                          .join("")}
                      </select>
                    </div>
                  </div>
                </div>
                <div class="field">
                  <label class="label" for="listNominativeTable" data-tooltip="Choisissez le tableau contenant vos données principales">Tableau :</label>
                  <div class="control">
                    <div class="select is-fullwidth">
                      <select name="listNominativeTable" id="listNominativeTable">
                        ${self.listNominativeTables
                          .map(
                            (table) =>
                              `<option value="${table.name}" ${
                                table.name === self.listNominativeTable ? "selected" : ""
                              }>${table.name}</option>`
                          )
                          .join("")}
                      </select>
                    </div>
                  </div>
                </div>
                <div class="field">
                  <label class="label" for="listNominativeJoinColumn" data-tooltip="Sélectionnez la colonne pour joindre avec la Liste Effective">Colonne de Jointure :</label>
                  <div class="control">
                    <div class="select is-fullwidth">
                      <select name="listNominativeJoinColumn" id="listNominativeJoinColumn">
                        ${self.listNominativeColumns
                          .map(
                            (column) =>
                              `<option value="${column.name}" ${
                                column.name === self.listNominativeJoinColumn ? "selected" : ""
                              }>${column.name}</option>`
                          )
                          .join("")}
                      </select>
                    </div>
                  </div>
                </div>
                <div class="field">
                  <div class="accordion">
                    <div class="accordion-header" onclick="app.toggleAdvanced('nominative')">
                      <span class="button is-small is-fullwidth grid" data-tooltip="Personnalisez les colonnes à inclure depuis la Liste Nominative">
                        <span class="icon"><i class="fas ${
                          self.showNominativeAdvanced ? "fa-angle-down" : "fa-angle-right"
                        }"></i></span> Modifier la Sélection des Colonnes
                      </span>
                    </div>
                    <div class="accordion-content ${self.showNominativeAdvanced ? "is-active" : ""}">
                      <label class="label">Colonnes à Inclure :</label>
                      <div class="buttons mb-2">
                        <button class="button is-small is-info" id="nominativeColumns-select-all" data-tooltip="Sélectionner toutes les colonnes disponibles">Tout Sélectionner</button>
                        <button class="button is-small is-warning" id="nominativeColumns-unselect-all" data-tooltip="Désélectionner toutes les colonnes">Tout Désélectionner</button>
                      </div>
                      <div id="nominativeColumns" class="grid is-mobile is-gap-0"></div>
                    </div>
                  </div>
                </div>
              </div>
              <div class="column">
                <h2 class="subtitle is-5 has-text-weight-bold">Liste Effective</h2>
                <div class="field">
                  <label class="label" for="listEffetiveSheet" data-tooltip="Sélectionnez la feuille contenant votre tableau de données secondaire">Feuille :</label>
                  <div class="control">
                    <div class="select is-fullwidth">
                      <select name="listEffetiveSheet" id="listEffetiveSheet">
                        ${self.sheets
                          .map(
                            (sheet) =>
                              `<option value="${sheet.name}" ${
                                sheet.name === self.listEffetiveSheet ? "selected" : ""
                              }>${sheet.name}</option>`
                          )
                          .join("")}
                      </select>
                    </div>
                  </div>
                </div>
                <div class="field">
                  <label class="label" for="listEffetiveTable" data-tooltip="Choisissez le tableau contenant vos données secondaires">Tableau :</label>
                  <div class="control">
                    <div class="select is-fullwidth">
                      <select name="listEffetiveTable" id="listEffetiveTable">
                        ${self.listEffetiveTables
                          .map(
                            (table) =>
                              `<option value="${table.name}" ${
                                table.name === self.listEffetiveTable ? "selected" : ""
                              }>${table.name}</option>`
                          )
                          .join("")}
                      </select>
                    </div>
                  </div>
                </div>
                <div class="field">
                  <label class="label" for="listEffetiveJoinColumn" data-tooltip="Sélectionnez la colonne pour joindre avec la Liste Nominative">Colonne de Jointure :</label>
                  <div class="control">
                    <div class="select is-fullwidth">
                      <select name="listEffetiveJoinColumn" id="listEffetiveJoinColumn">
                        ${self.listEffetiveColumns
                          .map(
                            (column) =>
                              `<option value="${column.name}" ${
                                column.name === self.listEffetiveJoinColumn ? "selected" : ""
                              }>${column.name}</option>`
                          )
                          .join("")}
                      </select>
                    </div>
                  </div>
                </div>
                <div class="field">
                  <div class="accordion">
                    <div class="accordion-header" onclick="app.toggleAdvanced('effective')">
                      <span class="button is-small is-fullwidth grid" data-tooltip="Personnalisez les colonnes à inclure depuis la Liste Effective">
                        <span class="icon"><i class="fas ${
                          self.showEffectiveAdvanced ? "fa-angle-down" : "fa-angle-right"
                        }"></i></span> Modifier la Sélection des Colonnes
                      </span>
                    </div>
                    <div class="accordion-content ${self.showEffectiveAdvanced ? "is-active" : ""}">
                      <label class="label">Colonnes à Inclure :</label>
                      <div class="buttons mb-2">
                        <button class="button is-small is-info" id="effectiveColumns-select-all" data-tooltip="Sélectionner toutes les colonnes disponibles">Tout Sélectionner</button>
                        <button class="button is-small is-warning" id="effectiveColumns-unselect-all" data-tooltip="Désélectionner toutes les colonnes">Tout Désélectionner</button>
                      </div>
                      <div id="effectiveColumns" class="grid is-mobile is-gap-0"></div>
                    </div>
                  </div>
                </div>
              </div>
            </div>
            <div class="field">
              <div class="accordion">
                <div class="accordion-header" onclick="app.toggleGroupBy()">
                  <span class="button is-fullwidth is-small grid" data-tooltip="Choisissez les colonnes pour regrouper les tableaux de sortie">
                    <span class="icon"><i class="fas ${
                      self.showGroupBy ? "fa-angle-down" : "fa-angle-right"
                    }"></i></span>
                    <span class="group-by-summary">Colonnes de Regroupement${groupBySummary}</span>
                  </span>
                </div>
                <div class="accordion-content ${self.showGroupBy ? "is-active" : ""}">
                  <label class="label">Sélectionnez les Colonnes pour Regrouper :</label>
                  <div id="nominativeGroupByColumns" class="grid is-mobile is-gap-0 is-columns-gap-0"></div>
                </div>
              </div>
            </div>
            <div class="field">
              <div class="accordion">
                <div class="accordion-header" onclick="app.toggleExtensionTables()">
                  <span class="button is-fullwidth is-small grid" data-tooltip="Ajoutez des tableaux supplémentaires pour enrichir vos données">
                    <span class="icon"><i class="fas ${
                      self.showExtensionTables ? "fa-angle-down" : "fa-angle-right"
                    }"></i></span>
                    <span>Tableaux d'Extension (${self.extensionTables.length})</span>
                  </span>
                </div>
                <div class="accordion-content ${self.showExtensionTables ? "is-active" : ""}">
                  <div id="extensionTablesContainer">
                    ${self.extensionTables
                      .map(
                        (ext, index) => `
                    <div class="box mb-2">
                      <h3 class="subtitle is-6">Tableau d'Extension ${index + 1}</h3>
                      <div class="field">
                        <label class="label" data-tooltip="Sélectionnez la feuille pour ce tableau d'extension">Feuille :</label>
                        <div class="control">
                          <div class="select is-fullwidth">
                            <select class="extension EFFECTSSheet" data-index="${index}">
                              ${self.sheets
                                .map(
                                  (s) =>
                                    `<option value="${s.name}" ${s.name === ext.sheet ? "selected" : ""}>${
                                      s.name
                                    }</option>`
                                )
                                .join("")}
                            </select>
                          </div>
                        </div>
                      </div>
                      <div class="field">
                        <label class="label" data-tooltip="Choisissez le tableau pour enrichir vos données">Tableau :</label>
                        <div class="control">
                          <div class="select is-fullwidth">
                            <select class="extensionTable" data-index="${index}">
                              ${ext.tables
                                .map(
                                  (t) =>
                                    `<option value="${t.name}" ${t.name === ext.table ? "selected" : ""}>${
                                      t.name
                                    }</option>`
                                )
                                .join("")}
                            </select>
                          </div>
                        </div>
                      </div>
                      <div class="field">
                        <label class="label" data-tooltip="Colonne du résultat pour joindre avec ce tableau">Colonne de Jointure (Résultat) :</label>
                        <div class="control">
                          <div class="select is-fullwidth">
                            <select class="extensionJoinResult" data-index="${index}">
                              ${self.orderedColumns
                                .map(
                                  (col) =>
                                    `<option value="${col}" ${
                                      col === ext.joinColumnResult ? "selected" : ""
                                    }>${col}</option>`
                                )
                                .join("")}
                            </select>
                          </div>
                        </div>
                      </div>
                      <div class="field">
                        <label class="label" data-tooltip="Colonne de ce tableau pour joindre avec le résultat">Colonne de Jointure (Extension) :</label>
                        <div class="control">
                          <div class="select is-fullwidth">
                            <select class="extensionJoinExt" data-index="${index}">
                              ${ext.columns
                                .map(
                                  (col) =>
                                    `<option value="${col.name}" ${
                                      col.name === ext.joinColumnExtension ? "selected" : ""
                                    }>${col.name}</option>`
                                )
                                .join("")}
                            </select>
                          </div>
                        </div>
                      </div>
                      <div class="field">
                        <div class="accordion">
                          <div class="accordion-header" onclick="app.toggleExtensionColumns(${index})">
                            <span class="button is-small is-fullwidth grid" data-tooltip="Sélectionnez les colonnes à inclure depuis ce tableau">
                              <span class="icon"><i class="fas ${
                                ext.showColumns ? "fa-angle-down" : "fa-angle-right"
                              }"></i></span> Modifier la Sélection des Colonnes
                            </span>
                          </div>
                          <div class="accordion-content ${ext.showColumns ? "is-active" : ""}">
                            <label class="label">Colonnes à Inclure :</label>
                            <div class="buttons mb-2">
                              <button class="button is-small is-info" id="extensionColumns-${index}-select-all" data-tooltip="Sélectionner toutes les colonnes">Tout Sélectionner</button>
                              <button class="button is-small is-warning" id="extensionColumns-${index}-unselect-all" data-tooltip="Désélectionner toutes les colonnes">Tout Désélectionner</button>
                            </div>
                            <div id="extensionColumns-${index}" class="grid is-mobile is-gap-0"></div>
                          </div>
                        </div>
                      </div>
                      <button class="button is-danger is-small removeExtension" data-index="${index}" data-tooltip="Supprimer ce tableau d'extension">Supprimer</button>
                    </div>
                  `
                      )
                      .join("")}
                  </div>
                  <button class="button is-info is-small" id="addExtensionTable" data-tooltip="Ajouter un nouveau tableau d'extension">Ajouter un Tableau d'Extension</button>
                </div>
              </div>
            </div>
            <div class="field">
              <div class="accordion">
                <div class="accordion-header" onclick="app.toggleColumnOrdering()">
                  <span class="button is-fullwidth is-small grid" data-tooltip="Réorganiser l'ordre des colonnes dans la sortie">
                    <span class="icon"><i class="fas ${
                      self.showColumnOrdering ? "fa-angle-down" : "fa-angle-right"
                    }"></i></span> Réorganiser les Colonnes
                  </span>
                </div>
                <div class="accordion-content ${self.showColumnOrdering ? "is-active" : ""}">
                  <div id="columnOrderList" class="box mb-2">
                    ${self.orderedColumns
                      .map(
                        (col) => `
                      <div class="column-order-item" draggable="true" data-column="${col}">
                        <span class="drag-handle">☰</span> ${col}
                      </div>
                    `
                      )
                      .join("")}
                  </div>
                  <button class="button is-small is-info" onclick="app.resetColumnOrder()" data-tooltip="Réinitialiser à l'ordre par défaut des colonnes">Réinitialiser à l'Ordre par Défaut</button>
                </div>
              </div>
            </div>
            <div class="field">
              <label class="label" data-tooltip="Filtrer les données de la Liste Nominative en fonction des valeurs des colonnes">Filtres :</label>
              <div id="filtersContainer">
                ${self.filters
                  .map((filter, index) => {
                    const uniqueValues = self.columnUniqueValues.get(filter.column) || [];
                    return `
                    <div class="field is-grouped mb-2">
                      <div class="control">
                        <div class="select">
                          <select class="filterColumn" data-index="${index}" data-tooltip="Choisissez une colonne à filtrer">
                            ${self.listNominativeColumns
                              .map(
                                (column) =>
                                  `<option value="${column.name}" ${column.name === filter.column ? "selected" : ""}>${
                                    column.name
                                  }</option>`
                              )
                              .join("")}
                          </select>
                        </div>
                      </div>
                      <div class="control">
                        <div class="select">
                          <select class="filterValue" data-index="${index}" data-tooltip="Sélectionnez une valeur pour filtrer la colonne">
                            <option value="">Sélectionnez une valeur</option>
                            ${uniqueValues
                              .map(
                                (value) =>
                                  `<option value="${value}" ${
                                    value === filter.value ? "selected" : ""
                                  }>${value}</option>`
                              )
                              .join("")}
                          </select>
                        </div>
                      </div>
                      <div class="control">
                        <button class="button is-danger removeFilter" data-index="${index}" data-tooltip="Supprimer ce filtre">Supprimer</button>
                      </div>
                    </div>
                  `;
                  })
                  .join("")}
              </div>
              <button class="button is-info mt-2 is-small" id="addFilter" data-tooltip="Ajouter un nouveau filtre pour affiner vos données">
                <span class="icon"><i class="fas fa-filter"></i></span> <span>Ajouter un Filtre</span>
              </button>
            </div>
            <div class="field grid">
              <div>
                <label class="label" for="templateSheet" data-tooltip="Sélectionnez la feuille contenant votre modèle d'en-tête">Feuille du Modèle d'En-tête :</label>
                <div class="control">
                  <div class="select is-fullwidth">
                    <select id="templateSheet">
                      ${self.allSheets
                        .map(
                          (sheet) =>
                            `<option value="${sheet.name}" ${sheet.name === self.templateSheet ? "selected" : ""}>${
                              sheet.name
                            }</option>`
                        )
                        .join("")}
                    </select>
                  </div>
                </div>
              </div>
              <div>
                <label class="label" for="templateRangeAddress" data-tooltip="Spécifiez la plage (ex. A1:B2) pour le modèle d'en-tête">Plage du Modèle (ex. A1:B2) :</label>
                <div class="control">
                  <input class="input" type="text" id="templateRangeAddress" placeholder="ex. A1:B2" value="${
                    self.templateRangeAddress
                  }">
                </div>
                <p class="help">Détectée automatiquement à partir de la plage utilisée de la feuille sélectionnée. Remplacez par une plage personnalisée si nécessaire.</p>
              </div>
            </div>
            <div class="field">
              <label class="label" for="resultSheetName" data-tooltip="Nommez la feuille où les données traitées seront sorties">Nom de la Feuille de Résultat :</label>
              <div class="control">
                <input class="input" type="text" id="resultSheetName" placeholder="Entrez le nom de la feuille de résultat" value="${
                  self.resultSheetName
                }">
              </div>
            </div>
            <div class="mt-4">
              <button class="button is-primary is-fullwidth" onclick="app.process()" data-tooltip="Démarrer le traitement des données selon vos sélections">Traiter les Données</button>
            </div>
          </div>
        `;

        if (self.showColumnOrdering) {
          self.setupDragAndDrop();
        }

        document.getElementById("listNominativeSheet")?.addEventListener(
          "change",
          self.debounce((e) => {
            self.listNominativeSheet = e.target.value;
            self.handleNominativeSheetChange();
          }, DEBOUNCE_DELAY)
        );
        document.getElementById("listNominativeTable")?.addEventListener(
          "change",
          self.debounce((e) => {
            self.listNominativeTable = e.target.value;
            self.handleNominativeTableChange();
          }, DEBOUNCE_DELAY)
        );
        document.getElementById("listNominativeJoinColumn")?.addEventListener(
          "change",
          self.debounce((e) => {
            self.listNominativeJoinColumn = e.target.value;
            self.handleNominativeJoinColumnChange();
          }, DEBOUNCE_DELAY)
        );
        document.getElementById("listEffetiveSheet")?.addEventListener(
          "change",
          self.debounce((e) => {
            self.listEffetiveSheet = e.target.value;
            self.handleEffectiveSheetChange();
          }, DEBOUNCE_DELAY)
        );
        document.getElementById("listEffetiveTable")?.addEventListener(
          "change",
          self.debounce((e) => {
            self.listEffetiveTable = e.target.value;
            self.handleEffectiveTableChange();
          }, DEBOUNCE_DELAY)
        );
        document.getElementById("listEffetiveJoinColumn")?.addEventListener(
          "change",
          self.debounce((e) => {
            self.listEffetiveJoinColumn = e.target.value;
            self.handleEffectiveJoinColumnChange();
          }, DEBOUNCE_DELAY)
        );
        document.getElementById("resultSheetName")?.addEventListener("input", (e) => {
          self.resultSheetName = self.sanitizeSheetName(e.target.value.trim() || DEFAULT_RESULT_SHEET);
        });
        document.getElementById("templateSheet")?.addEventListener(
          "change",
          self.debounce(async (e) => {
            self.templateSheet = e.target.value;
            await self.updateTemplateRange(self.templateSheet);
          }, DEBOUNCE_DELAY)
        );
        document.getElementById("templateRangeAddress")?.addEventListener(
          "input",
          self.debounce((e) => {
            self.templateRangeAddress = e.target.value.trim();
          }, DEBOUNCE_DELAY)
        );

        document.getElementById("addFilter")?.addEventListener("click", async () => {
          if (self.listNominativeColumns.length > 0) {
            const newColumn = self.listNominativeColumns[0].name;
            self.filters.push({ column: newColumn, value: "" });
            await self.getUniqueColumnValues(newColumn);
            self.render();
          }
        });

        document.querySelectorAll(".filterColumn").forEach((select) => {
          select.addEventListener(
            "change",
            self.debounce(async (e) => {
              const index = parseInt(e.target.dataset.index, 10);
              self.filters[index].column = e.target.value;
              self.filters[index].value = "";
              if (!self.columnUniqueValues.has(self.filters[index].column)) {
                await self.getUniqueColumnValues(self.filters[index].column);
              }
              self.render();
            }, DEBOUNCE_DELAY)
          );
        });

        document.querySelectorAll(".filterValue").forEach((select) => {
          select.addEventListener(
            "change",
            self.debounce((e) => {
              const index = parseInt(e.target.dataset.index, 10);
              self.filters[index].value = e.target.value;
            }, DEBOUNCE_DELAY)
          );
        });

        document.querySelectorAll(".removeFilter").forEach((button) => {
          button.addEventListener(
            "click",
            self.debounce((e) => {
              const index = parseInt(e.target.dataset.index, 10);
              self.filters.splice(index, 1);
              self.render();
            }, DEBOUNCE_DELAY)
          );
        });

        document.getElementById("addExtensionTable")?.addEventListener("click", () => {
          self.addExtensionTable();
        });
        document.querySelectorAll(".extensionSheet").forEach((select) => {
          select.addEventListener(
            "change",
            self.debounce((e) => {
              const index = parseInt(e.target.dataset.index);
              self.extensionTables[index].sheet = e.target.value;
              self.handleExtensionSheetChange(index);
            }, DEBOUNCE_DELAY)
          );
        });
        document.querySelectorAll(".extensionTable").forEach((select) => {
          select.addEventListener(
            "change",
            self.debounce((e) => {
              const index = parseInt(e.target.dataset.index);
              self.extensionTables[index].table = e.target.value;
              self.handleExtensionTableChange(index);
            }, DEBOUNCE_DELAY)
          );
        });
        document.querySelectorAll(".extensionJoinResult").forEach((select) => {
          select.addEventListener(
            "change",
            self.debounce((e) => {
              const index = parseInt(e.target.dataset.index);
              self.extensionTables[index].joinColumnResult = e.target.value;
              self.render();
            }, DEBOUNCE_DELAY)
          );
        });
        document.querySelectorAll(".extensionJoinExt").forEach((select) => {
          select.addEventListener(
            "change",
            self.debounce((e) => {
              const index = parseInt(e.target.dataset.index);
              self.extensionTables[index].joinColumnExtension = e.target.value;
              self.render();
            }, DEBOUNCE_DELAY)
          );
        });
        document.querySelectorAll(".removeExtension").forEach((button) => {
          button.addEventListener(
            "click",
            self.debounce((e) => {
              const index = parseInt(e.target.dataset.index);
              self.removeExtensionTable(index);
            }, DEBOUNCE_DELAY)
          );
        });

        self.updateDropdowns();
        self.extensionTables.forEach((ext, index) => {
          self.updateColumnSelection(`extensionColumns-${index}`, ext.columns, ext.selectedColumns);
        });
      } catch (error) {
        self.log("error", "Erreur de rendu de l'interface utilisateur:", error);
        self.showError("Échec du rendu de l'interface utilisateur. Veuillez actualiser le complément.");
      }
    };

    console.log("before init");
    
    self
      .init()
      .then(() => {
        console.log("befor");
        
        self.render();
        console.log("rendered");
        



      })
      .catch((error) => {
        console.log("init error +++");
        
        self.log("error", "Échec de l'initialisation:", error);
        self.showError(`Échec de l'initialisation : ${error.message}`);
      });
  }


  // console.log("hilaaaaw");

  
 const app = new App();
  window.app = app;
});
