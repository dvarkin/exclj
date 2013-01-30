(ns exclj.math-xls-test
  (:use clojure.test
        exclj.math_xls))

(deftest cdf-test
  (testing "CDF calc is incorrect. Excel version calc this as 0.36944134"
    (is (= 0.36944134018176367 (normal-distribution 10 20 30 true)))))

(deftest pdf-test
  (testing "CDF calc is incorrect. Excel version calc this as 0.012579441"
    (is (= 0.012579440923099773 (normal-distribution 10 20 30 false)))))
