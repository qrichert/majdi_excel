test:
	@python main.py tests/test_a.csv tests/test_b.csv

test_extra_column:
	@python main.py tests/test_a.csv tests/test_extra_column.csv
