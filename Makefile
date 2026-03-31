.PHONY: help install install-dev sync test test-unit test-e2e test-cov lint format format-check type-check check clean run server

help:
	@echo "Available targets:"
	@echo "  make install          - Install project dependencies"
	@echo "  make install-dev      - Install project with dev dependencies"
	@echo "  make test             - Run all unit tests (excluding e2e)"
	@echo "  make test-cov         - Run tests with coverage report"
	@echo "  make lint             - Run ruff linter"
	@echo "  make format           - Format code with ruff"
	@echo "  make check            - Run all quality checks"
	@echo "  make clean            - Clean build artifacts"
	@echo "  make run              - Run the MCP server"

install:
	uv sync

install-dev:
	uv sync --group dev --group test

sync:
	uv sync --group dev --group test

test:
	uv run pytest tests/ -m "not e2e" -v

test-unit:
	uv run pytest tests/ -m "not e2e" -v

test-e2e:
	ONEDRIVE_E2E=1 uv run pytest tests/e2e/ -v -m e2e

test-cov:
	uv run pytest tests/ -m "not e2e" --cov=src/onedrive_blade_mcp --cov-report=term-missing -v

lint:
	uv run ruff check src/ tests/

format:
	uv run ruff format src/ tests/

format-check:
	uv run ruff format --check src/ tests/

type-check:
	uv run mypy src/onedrive_blade_mcp

check: lint format-check type-check
	@echo "All quality checks passed!"

run:
	uv run onedrive-blade-mcp

server: run

clean:
	rm -rf .pytest_cache .mypy_cache .ruff_cache htmlcov dist build *.egg-info
	find . -type d -name __pycache__ -exec rm -rf {} + 2>/dev/null || true
	find . -type f -name "*.pyc" -delete
	@echo "Cleanup complete!"
