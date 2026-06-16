COMPOSE := docker compose
SERVICE := zoho-data
PYTHON := .venv/bin/python

.PHONY: help up down restart rebuild logs status shell dev clean

help:
	@printf "%s\n" \
	"Common commands:" \
	"  make up       Start the app in Docker Compose background mode" \
	"  make down     Stop the app and remove the container" \
	"  make restart  Restart the running app container" \
	"  make rebuild  Rebuild the image and restart the app" \
	"  make logs     Follow application logs" \
	"  make status   Show container status" \
	"  make shell    Open a shell inside the running container" \
	"  make dev      Run the app locally with the project virtualenv" \
	"  make clean    Remove disposable Docker artifacts for this app only"

up:
	$(COMPOSE) up -d

down:
	$(COMPOSE) down

restart:
	$(COMPOSE) restart

rebuild:
	$(COMPOSE) up -d --build

logs:
	$(COMPOSE) logs -f $(SERVICE)

status:
	$(COMPOSE) ps

shell:
	$(COMPOSE) exec $(SERVICE) /bin/sh

dev:
	$(PYTHON) app/server.py

clean:
	$(COMPOSE) down --remove-orphans --rmi local
