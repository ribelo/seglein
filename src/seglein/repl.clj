(ns seglein.repl
  (:require [clojure.java.io :as io]
            [clojure.data.csv :as csv]
            [dk.ative.docjure.spreadsheet :as xls]
            [cuerdas.core :as str]
            [clj-pdf.core :as pdf]))


(def group-map (->> (xls/load-workbook-from-file "/home/huxley/schowek/grupy.xls")
                    (xls/select-sheet "Arkusz 1")
                    (xls/select-columns {:A :group/name
                                         :B :group/index})
                    (rest)
                    (drop-last 4)
                    (mapv (fn [{:keys [group/name group/index] :as m}] {index name}))
                    (into {})))

(-> group-map (first))

(def input-data
  (->> (xls/load-workbook-from-file "/home/huxley/schowek/obroty.xls")
       (xls/select-sheet "Arkusz 1")
       (xls/select-columns {:A :group/index
                            :B :name
                            :C :index
                            :D :margin
                            :E :sales
                            :F :purchases
                            :G :promotion})
       (rest)
       (mapv (fn [m]
               (as-> m $
                   (update $ :name (fn [s] (-> s (str/replace "\"" "") (str/clean))))
                   (update $ :group/index (fn [s] (apply str (take 4 s))))
                   (assoc $ :group/name (let [{:keys [group/index]} $] (get group-map index))))))))

(def output-data
  (->> input-data
       (group-by :group/index)
       (map (fn [[index coll]]
              (let [sales-all (->> coll (transduce (map :sales) +))
                    sales-promotion (->> coll (transduce (comp (remove #(#{" "} (:promotion %))) (map :sales)) +))
                    sales-non-promotion (- sales-all sales-promotion)
                    purchases-all (->> coll (transduce (map :purchases) +))
                    purchases-promotion (->> coll (transduce (comp (remove #(#{" "} (:promotion %))) (map :purchases)) +))
                    purchases-non-promotion (- purchases-all purchases-promotion)
                    margin-all-value (- sales-all purchases-all)
                    margin-promotion-value (- sales-promotion purchases-promotion)
                    margin-non-promotion-value (- sales-non-promotion purchases-non-promotion)
                    margin-all (if (pos? sales-all) (/ margin-all-value sales-all) 0.0)
                    margin-promotion (if (pos? sales-promotion) (/ margin-promotion-value sales-promotion) 0.0)
                    margin-non-promotion (if (pos? sales-non-promotion) (/ margin-non-promotion-value sales-non-promotion) 0.0)
                    target-margin (->> coll (map :margin) (frequencies) (ffirst) (* 0.01))
                    target-margin-value (- (/ purchases-all (- 1.0 target-margin)) purchases-all)
                    sugestet-margin (cond
                                      (zero? sales-non-promotion) "brak sprzedaży poza promocją"
                                      (zero? sales-promotion) "brak sprzedaży w promocji"
                                      (> margin-promotion target-margin) "promocyjna marża wieksza od założonej"
                                      (pos? sales-non-promotion) (/ (- target-margin-value margin-promotion-value) sales-non-promotion))]
                {:group/index          index
                 :group/name           (->> coll (first) :group/name)
                 ;:sales/all               sales-all
                 ;:sales/promotion         sales-promotion
                 ;:sales/non-promotion     sales-non-promotion
                 ;:purchases/all           purchases-all
                 ;:purchases/promotion     purchases-promotion
                 ;:purchases/non-promotion purchases-non-promotion
                 :margin/all           margin-all
                 :margin/promotion     margin-promotion
                 :margin/non-promotion margin-non-promotion
                 ;:margin/promotion-value  margin-promotion-value
                 ;:margin/target-value     target-margin-value
                 :margin/target        target-margin
                 :margin/sugested      sugestet-margin
                 ;:diff                    (- target-margin-value margin-promotion-value)
                 })))
       (sort-by :group/index)))

(pdf/pdf
  [{:title                  "kalkulacja marży"
    :right-margin           50
    :left-margin            10
    :bottom-margin          10
    :top-margin             10
    :size                   "a4"
    :register-system-fonts? true
    :font                   {:encoding :unicode
                             :align    :center
                             :size     10
                             :ttf-name "Noto Sans"}}
   [:paragraph {:style :bold :size 14}
    "kalkulacja marży"]
   [:spacer 3]
   (into [:table {:header ["index" "nazwa" "marża założona" "mraża całość"
                           "marża z promocją" "marża poza promocją"
                           "sugerowana marża"]}]
         (for [{:keys [group/index group/name margin/target
                       margin/all margin/promotion
                       margin/non-promotion margin/sugested]} output-data]
           [[:cell index]
            [:cell (or name "błąd")]
            [:cell (format "%.2f" target)]
            [:cell (format "%.2f" all)]
            [:cell (if-not (zero? promotion) (format "%.2f" promotion) "brak")]
            [:cell (if-not (zero? non-promotion) (format "%.2f" non-promotion) "brak")]
            [:cell (if (number? sugested) (format "%.2f" sugested) sugested)]
            ]))]
  "/home/huxley/schowek/kalkulacja-marzy-sklep-1.pdf")