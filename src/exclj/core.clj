(ns exclj.core
  (:use
   [clojure.set :only (index)]
   [exclj.xls_types :only (extract pack)]
   [exclj.xls_read :only (read-workbook map-workbook)]
   [exclj.outcome :only (parse-outcomes get-outcomes)]
   [exclj.params :only (parse-params set-params get-params)]
   )
  (:require [clojure.java.io :as io])
  (:import [java.io PushbackReader]))


(defrecord WorkbookConfig
    [id wb params sheets in-sheet out-sheet file-path])

;;;; map of "file-path" = parsed math model.
(def workbooks (ref nil))

;;;; read config from $PROJ_ROOT/config file
(def config (with-open [r (io/reader "config")]
            (read (PushbackReader. r))))

(defn config-to-workbook [config]
  (let [id (config :id)
        file-path (config :file-path)
        wb (read-workbook file-path)
        sheets (map-workbook wb)
        in-sheet (sheets (config :in-sheet-name))
        out-sheet (sheets (config :out-sheet-name))
        params (index (parse-params  in-sheet) [:id])]
    (->WorkbookConfig id wb params sheets in-sheet out-sheet file-path)))

(defn calc [params]
  (let [id (params :id)
        config (first (filter #(= (:id %) id) @workbooks))
        new-config (set-params config params)]
    (parse-outcomes (:out-sheet new-config))))
;    (println (get-params [10 11] new-config))
;    (get-outcomes [1 2] new-config)))
;    (parse-outcomes (:out-sheet new-config))))

(defn load-models []
  ;; config from global config
  (let [models (:math_models config)]
    (map config-to-workbook models)))

(defn boot []
  (dosync (ref-set workbooks (load-models))))

(defn -main
  [& args]
  boot)
                   
