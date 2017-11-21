(ns seglein.repl
  (:require [clojure.java.io :as io]
            [clojure.data.csv :as csv]
            [dk.ative.docjure.spreadsheet :as xls]
            [cuerdas.core :as str]))


(def group-map (->> (xls/load-workbook-from-file "/home/huxley/schowek/grupy.xls")
                    (xls/select-sheet "Arkusz 1")
                    (xls/select-columns {:A :group/name
                                         :B :group/index})
                    (rest)
                    (drop-last 4)
                    (mapv (fn [{:keys [group/name group/index] :as m}] {index name}))
                    (into {})))

(-> group-map (first))

(def data
  (->> (xls/load-workbook-from-file "/home/huxley/schowek/obroty.xls")
       (xls/select-sheet "Arkusz 1")
       (xls/select-columns {:A :group/index
                            :B :name
                            :C :index
                            :D :sales
                            :E :purchases
                            :F :promotion})
       (rest)
       (mapv (fn [m]
               (-> m
                   (update :name (fn [s] (-> s (str/replace "\"" "") (str/clean))))
                   (assoc :group/name (let [{:keys [group/index]} m] (get group-map index))))))))



(->> data
     (first))
(->> data
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
                  target-margin 0.2
                  target-margin-value (- (/ purchases-all (- 1.0 target-margin)) purchases-all)
                  sugestet-margin (cond
                                    (zero? sales-non-promotion) "brak sprzedaży poza promocją"
                                    (> margin-promotion target-margin) "promocyjna marża wieksza od założonej"
                                    (pos? sales-non-promotion) (/ (- target-margin-value margin-promotion-value) sales-non-promotion) )]
              {:group/index             index
               :group/name              (->> coll (first) :group/name)
               :sales/all               sales-all
               :sales/promotion         sales-promotion
               :sales/non-promotion     sales-non-promotion
               :purchases/all           purchases-all
               :purchases/promotion     purchases-promotion
               :purchases/non-promotion purchases-non-promotion
               :margin/all              margin-all
               :margin/promotion        margin-promotion
               :margin/non-promotion    margin-non-promotion
               :margin/promotion-value  margin-promotion-value
               :margin/target-value     target-margin-value
               :margin/sugestet         sugestet-margin
               :diff                    (- target-margin-value margin-promotion-value)
               }))))