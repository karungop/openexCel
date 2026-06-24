import pathlib, pytest

FIXTURES = pathlib.Path(__file__).parent / "fixtures"

def fixture_path(name):
    p = FIXTURES / name
    if not p.exists():
        pytest.skip(f"Fixture {name} not found — run tests/make_fixtures.py first")
    return str(p)
