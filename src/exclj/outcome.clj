(ns exclj.outcome
  (:use [exclj.xls_types :only (CastClass)]
   [exclj.xls_read :only (values rows cells parse-sheet)]))

(defrecord Outcome
  [id market-type market-name  name coef param]
  CastClass
  (toString [this] (str id " "  market-type " "  market-name " " param " "  name " " coef "\n")))

(defn init-outcome
  [[id market-type market-name param name coef]]
  (->Outcome (int id) market-type market-name param name coef))

(defn parse-outcomes
  "parse and calc outcomes from outcomes sheet OUT - by default"
  [sheet]
  (parse-sheet sheet #(init-outcome (vec (values %)))))

(defn get-outcomes
  "get outcome by id"
  [ids config]
  (->> (:out-sheet config)
       parse-outcomes
       (filter #(some (fn [x] (= (:id %) x)) ids))))
