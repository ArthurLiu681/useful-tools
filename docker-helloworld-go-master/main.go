package main

import (
	"fmt"
	"log"
	"net/http"
	"os"
)

const PORT_ENV_VAR = "PORT"
const PORT_DEFAULT = "8080"

func main() {
	http.HandleFunc("/", func(w http.ResponseWriter, r *http.Request) {
		msg := os.Getenv("SIMPLE_MSG")
	        if msg == "" {
		   msg = "Hello world!"
	        }
	        fmt.Fprintf(w, "<h1>%s</h1>", msg)
	})

	port := os.Getenv(PORT_ENV_VAR)
	if port == "" {
		port = PORT_DEFAULT
	}
	addr := ":" + port

	log.Printf("Listening on %s", addr)
	err := http.ListenAndServe(addr, nil)
	log.Fatalf("Error listening and serving on: %#v", err)
}
