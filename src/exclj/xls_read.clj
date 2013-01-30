(ns exclj.xls_read
  (:use
   [clojure.java.io :only [file input-stream]]
   [exclj.xls_types :only (extract pack)])
  (:import
   (org.apache.poi.ss.usermodel WorkbookFactory)
   (java.io File InputStream)))

(defn workbook
  "Create or open new excel workbook. Defaults to xlsx format."
  ([input] (WorkbookFactory/create input)))

(defn sheets
  "Get seq of sheets."
  [wb] (map #(.getSheetAt wb %1) (range 0 (.getNumberOfSheets wb))))

(defn rows
  "Return rows from sheet as seq.  Simple seq cast via Iterable implementation."
  [sheet] (seq sheet))

(defn cells
  "Return seq of cells from row.  Simpel seq cast via Iterable implementation." 
  [row] (seq row))

(defn get-address
  [cell]
  [(.getRowIndex cell)  (.getColumnIndex cell)])

(defn values
  "Return cells from sheet as seq."
  [row] (map extract (cells row)))

(defn values-address
  [row]
  (map #(conj (get-address %) (extract %)) (cells row)))

;; params

(defn parse-sheet
  "parse sheet and apply fn to every row"
  [sheet fn]
    (->> 
     (rows sheet)
     (map cells)
     (map fn)))

(defn read-workbook [file-path]
  (let [f (file file-path)
        input (input-stream f)]
    (with-open [in input]
      (workbook in))))

(defn map-workbook
  "Lazy workbook report."
  [wb]
  (zipmap (map #(.getSheetName %) (sheets wb)) (sheets wb)))

(defn get-cell
  "Sell cell within row"
  ([row col] (.getCell row col))
  ([sheet row col] (get-cell (or (.getRow sheet row) (.createRow sheet row)) col)))

(defn set-cell
  "set cell value"
  [sheet row col value]
  (let [cell (get-cell sheet row col)]
    (pack cell value)
    cell))




