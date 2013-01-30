(ns exclj.params
  (:use [exclj.xls_types :only (CastClass extract)]
        [exclj.xls_read  :only (get-cell set-cell values-address parse-sheet)]))

(defrecord Param
    [id type name value row col]
  CastClass
  (toString [this] (str id " " type " "  name " " value " " row " " col "\n")))

(defn init-param
  [[[_ _ id] [_ _ type] [_ _ name]  [row col value] ]]
  (->Param  (int id) type name value row col))

(defn parse-params
  "parse and calc outcomes from outcomes sheet OUT - by default"
  [sheet]
  (parse-sheet sheet #(init-param (values-address %))))

;;; functions for manipulate with config from core

(defn param-address [id params-records]
  ;;; extract [row col] of param from config
  (let [r (first (get params-records {:id id}))
        row (:row r) 
        col (:col r)]
    [row col]))

(defn get-params
  ;;; get value of param by id in Excel file
  [ids config]
  (let [params (:params config)
        sheet (:in-sheet config)
        address  (map #(let [[r c ] (param-address % params)]
                         (get-cell sheet r c))
                      ids)
        ]
    (map extract address)))

(defn set-param 
  ;;; set value of cell new is map with address of param cell
  [config new]
  (let [id (:id new)
        value (:value new)
        params (:params config)
        sheet (:in-sheet config)
        r (first (get params {:id id}))
        row (:row r) 
        col (:col r)
        ]
    (set-cell sheet row col value)
    (assoc config :in-sheet sheet)
    ))

(defn set-param [cell params])

;(def p {:params [{:id 10, :value 2} {:id 11, :value 3}], :id 1})
(defn set-params
  "take map of params {id:value}"
  [config params]
  (->>
   (:params params)
   (reduce #(set-param %1 %2) config)))


