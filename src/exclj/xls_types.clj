(ns exclj.xls_types
  (:use [exclj.math_xls :only [normal-distribution]])
  (:import (org.apache.poi.ss.formula.functions Fixed4ArgFunction)
           (org.apache.poi.ss.formula WorkbookEvaluator)
           (org.apache.poi.ss.formula.eval NumberEval BoolEval RefEvalBase)
           (org.apache.poi.ss.usermodel WorkbookFactory Cell CellValue)
           (org.apache.poi.hssf.usermodel HSSFCell)))

(defn formula-eval [cell]
  "eval excel formula"
  (let [evaluator (.. cell getSheet getWorkbook getCreationHelper createFormulaEvaluator)
        cv (.evaluate evaluator cell)]
    cv))

(defprotocol CastClass
  (extract [this] "cast to normall")
  (pack [this]  [this value] "turn clean class to some for Java")
  (toString [this] "print record"))

(extend-type NumberEval
  CastClass
  (extract [this] (.getNumberValue this)))

(extend-type java.lang.Double
    CastClass
    (pack [this] (NumberEval. this)))

(extend-type BoolEval
  CastClass
  (extract [this] (.getBooleanValue this)))

(extend-type  RefEvalBase
  CastClass
  (extract [this]
    (extract (.getInnerValueEval this))))

(extend-type HSSFCell
  CastClass
  (extract [this]
    (condp = (.getCellType this)
      Cell/CELL_TYPE_FORMULA    (extract (formula-eval this))
      Cell/CELL_TYPE_NUMERIC    (.getNumericCellValue this)
      Cell/CELL_TYPE_STRING     (.getStringCellValue this)
      Cell/CELL_TYPE_BOOLEAN    (.getBooleanCellValue this)
      Cell/CELL_TYPE_ERROR      (.getErrorCellValue this)
      Cell/CELL_TYPE_BLANK      nil
      (println "undef cell type: " (.getCellType this))))
  (pack [this value]
    (->> (double value) (.setCellValue this) )))

(extend-type Cell
  CastClass
  (extract [this]
    (condp = (.getCellType this)
      Cell/CELL_TYPE_FORMULA    (extract (formula-eval this))
      Cell/CELL_TYPE_NUMERIC    (.getNumericCellValue this)
      Cell/CELL_TYPE_STRING     (.getStringCellValue this)
      Cell/CELL_TYPE_BOOLEAN    (.getBooleanCellValue this)
      Cell/CELL_TYPE_ERROR      (.getErrorCellValue this)
      Cell/CELL_TYPE_BLANK      nil
      (println "undef cell type: " (.getCellType this))))
  (pack [this value] (.setCellValue this value)))

(extend-type CellValue
  CastClass
  (extract [this]
    (condp = (.getCellType this)
      Cell/CELL_TYPE_NUMERIC    (.getNumberValue this)
      Cell/CELL_TYPE_STRING     (.getStringValue this)
      Cell/CELL_TYPE_BOOLEAN    (.getBooleanValue this)
      Cell/CELL_TYPE_ERROR      (.getErrorValue this)
      (println "undef cell-value type: " (.getCellType this)))))

(def normdist
  (proxy [Fixed4ArgFunction] []
    (evaluate [col-index  row-index x mean standard_dev cumulative]
      (let [cumulative-value (extract cumulative)
            standard_dev-value (extract standard_dev)
            x-value (extract  x)
            mean-value (extract mean)
            ]
        (pack (normal-distribution x-value mean-value standard_dev-value cumulative-value))))))

(. WorkbookEvaluator registerFunction "NORMDIST" normdist)
