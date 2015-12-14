(ns table-to-b.core
  (:require [dk.ative.docjure.spreadsheet :as dj]
            [stencil.core :as sc]
            [clojure.string :as st])
  (:gen-class))

(defmacro forv [& body]
  `(into [] (for ~@body)))

(defn tuple-encoding [name table-name data]
  (sc/render-string
   "
MACHINE {{machine}}
 DEFINITIONS TUPLETYPE == ({{tuple-type}})
 ABSTRACT_CONSTANTS {{accessor-names}}
 CONSTANTS {{table-name}}

 PROPERTIES
  {{table-name}} : POW(TUPLETYPE) &
  {{table-name}} = {
   {{{tuples}}}
  } &
  {{{accessor-functions}}}
END
" (assoc data :machine name :table-name table-name)))

(defn record-encoding [name table-name data]
  (sc/render-string
   "
MACHINE {{machine}}
 DEFINITIONS STRUCTTYPE == ({{struct-type}})
 ABSTRACT_CONSTANTS {{accessor-names}}
 CONSTANTS {{table-name}}

 PROPERTIES
  {{table-name}} : POW(STRUCTTYPE) &
  {{table-name}} = {
   {{{records}}}
  } &
  {{{struct-functions}}}
END
" (assoc data :machine name :table-name table-name)))

(def types ["INTEGER" "STRING" :formula :blank "BOOL" :error])

(defn row-width [row] (count (dj/cell-seq row)))

(defn extract-table [cols rows]
  (forv [r rows]
        (forv [col (range cols)]
              (do
                (let [cell (.getCell r col)]
                  (when cell
                    (let [value (dj/read-cell cell)]
                      {:value (if (float? value) (int value) value)
                       :type (types (.getCellType cell))
                       :row (.getRowIndex cell)
                       :col (.getColumnIndex cell)})))))))

(defn fix-types [column]
  (let [etype (:type (first column))
        problems (remove #(= etype (:type %)) column)]
    (doseq [p problems]
      (binding [*out* *err*]
        (println "WARNING" (sc/render-string "Type mismatch in row {{row}} column {{col}}. Expected type is {{etype}} but value is of type {{type}}" (assoc p :etype etype)))))
    (if (seq problems)
      (map #(assoc % :type "STRING" :value (str (:value %))) column)
      column)))

(defn ascii [c]
  (let [i (int c)]
    (or (< 64 i 91) (< 96 i 123))))

(defn make-tuple [row]
  (let [elements (map (comp pr-str :value) row)]
    (str "(" (st/join "," elements) ")")))

(defn make-record [header row]
  (let [elements (map (comp pr-str :value) row)]
    (str "rec(" (st/join "," (map (fn [n v] (str n ":" v)) header elements)) ")")))

(defn make-accessor [n id name]
  (let [vars (st/join "," (for [x (range n)] (str "v" x)))]
    (sc/render-string
     "{{name}} = %({{vars}}).(({{vars}}):TUPLETYPE|v{{id}})"
     {:name name :vars vars :id id})))

(defn make-struct [name]
  (sc/render-string
   "{{name}} = %r.(r:STRUCTTYPE|r'{{name}})"
   {:name name }))

(defn parse-excel [filename sheetname]
  (let [workbook (dj/load-workbook filename)
        sheet (dj/select-sheet sheetname workbook)
        rows (dj/row-seq sheet)
        header (first rows)
        table (extract-table (count (dj/cell-seq header)) (rest rows))
        transposed (mapv fix-types (apply map vector table))
        table (apply map vector transposed)
        types (map (comp :type first) transposed)
        header (forv [c (dj/cell-seq header)] (apply str (filter ascii (dj/read-cell c))))]
    {:header header
     :table table
     :transposed transposed
     :types types
     :accessor-names (st/join "," header)
     :accessor-functions (st/join " &\n  " (map-indexed (partial make-accessor (count header)) header))
     :struct-functions (st/join " &\n  " (map make-struct header))
     :tuple-type (st/join "*" types)
     :struct-type (str "struct(" (st/join "," (map (fn [n t] (str n ":" t)) header types)) ")")
     :tuples (st/join ",\n   " (map make-tuple table))
     :records (st/join ",\n   " (map (partial make-record header) table))}))

(defn -main [& args]
  (let [[mode xls-file sheet table-name machine-name] args
        data (parse-excel xls-file sheet)
        output (cond (= mode "tuple") (tuple-encoding machine-name table-name data)
                     (= mode "record") (record-encoding machine-name table-name data))]
    (println output)))
