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
    Nummer string `json:"nummer"`
}

func suggestionsHandler(w http.ResponseWriter, r *http.Request) {
    q := r.URL.Query().Get("q")
    if len(q) < 1 {
        http.Error(w, "Query param 'q' min 1 Zeichen", http.StatusBadRequest)
        return
    }

    like := q + "%"
    rows, err := db.Query("SELECT nummer FROM tickets WHERE nummer LIKE ? ORDER BY nummer LIMIT 10", like)
    if err != nil {
        http.Error(w, "DB-Fehler", http.StatusInternalServerError)
        log.Println("DB query error:", err)
        return
    }
    defer rows.Close()

    var list []Ticket
    for rows.Next() {
        var t Ticket
        if err := rows.Scan(&t.Nummer); err != nil {
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
		dsn = "zabbix:testd@tcp(192.168.232.200:3306)/zabbix?parseTime=true"
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
