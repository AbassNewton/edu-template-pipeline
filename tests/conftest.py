import pytest

from edutpl.config import load_spec


@pytest.fixture(scope="session")
def spec():
    return load_spec()
