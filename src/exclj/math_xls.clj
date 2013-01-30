(ns exclj.math_xls
  (:import (org.apache.commons.math3.distribution NormalDistribution)))

(defn cdf
  "cumulatice version of Normal distribution"
  [Distirbution value]
  (.cumulativeProbability Distirbution value))

(defn pdf
  "probability version of Normal distribution"
  [Distirbution value]
  (.density Distirbution value))


(defn normal-distribution
  "function for calc normal distribution with PDF or CDF versions"
  [x mean standard_dev cumulative]
  (let [distribution-object  (NormalDistribution. mean standard_dev)]
    (if (true? cumulative)
      (cdf distribution-object x)
      (pdf distribution-object x))))
  
