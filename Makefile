.PHONY: test test-v install-test-deps

# Install pytest into the venv if not already present
install-test-deps:
	venv/bin/pip install pytest --quiet

# Run the full test suite
test: install-test-deps
	venv/bin/pytest tests/ -v --tb=short

# Run with verbose output and no capture (useful for debugging)
test-v: install-test-deps
	venv/bin/pytest tests/ -v --tb=long -s
