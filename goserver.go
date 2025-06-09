// main.go
package main

import (
	"database/sql"
	"encoding/json"
	"log"
	"net/http"
	"os"

	_ "github.com/go-sql-driver/mysql"
)

var db *sql.DB

type Ticket struct {
	Nummer string `json:"tn"`
	Title  string `json:"title"`
	Type   string `json:"name"`
}

func suggestionsHandler(w http.ResponseWriter, r *http.Request) {

	// Add CORS headers
	w.Header().Set("Access-Control-Allow-Origin", "https://mail.ulewu.de")
	w.Header().Set("Access-Control-Allow-Methods", "GET, OPTIONS")
	w.Header().Set("Access-Control-Allow-Headers", "Content-Type")

	// Handle preflight request
	if r.Method == http.MethodOptions {
		w.WriteHeader(http.StatusNoContent)
		return
	}

	q := r.URL.Query().Get("q")
	if len(q) < 1 {
		http.Error(w, "Query param 'q' min 1 Zeichen", http.StatusBadRequest)
		return
	}

	//like := q + "%"

	rows, err := db.Query("select tn, title, name  from ticket left join ticket_type on ticket.type_id = ticket_type.id LIMIT 10")
	if err != nil {
		http.Error(w, "DB-Fehler", http.StatusInternalServerError)
		log.Println("DB query error:", err)
		return
	}
	defer rows.Close()

	var list []Ticket
	for rows.Next() {
		var t Ticket
		if err := rows.Scan(&t.Nummer, &t.Title, &t.Type); err != nil {
			continue
		}
		list = append(list, t)
	}

	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(list)
}

func main() {
	var err error
	dsn := os.Getenv("MYSQL_DSN")
	if dsn == "" {
		// dsn = "user:password@tcp(localhost:3306)/deinedb?parseTime=true"
		dsn = "otobo:P351fpLqcS0gosk4@tcp(ncl1.chaos.local:3306)/otobo?parseTime=true"
	}
	db, err = sql.Open("mysql", dsn)
	if err != nil {
		log.Fatal(err)
	}
	defer db.Close()

	http.HandleFunc("/api/tickets/suggestions", suggestionsHandler)

	log.Println("Server läuft auf :8080")
	log.Fatal(http.ListenAndServe(":8080", nil))
}
