(defproject exclj "0.1.0-SNAPSHOT"
  :description "lib for read excel with formulas, and make output"
  :url "http://dvarkin.blogspot.com"
  :license {:name "Eclipse Public License"
            :url "http://www.eclipse.org/legal/epl-v10.html"}
  :dependencies [[org.clojure/clojure "1.4.0"]
                 [org.apache.poi/poi "3.8"]
                 [org.apache.poi/poi-ooxml "3.8"]
                 [org.apache.commons/commons-math3 "3.1.1"]
                 ]
  
  :plugins [[lein-swank "1.4.4"]]
  :aot [exclj.core]
  :main exclj.core
  :jvm-opts ["-Xmx1000m -Dfile.encoding=UTF8"]
)
