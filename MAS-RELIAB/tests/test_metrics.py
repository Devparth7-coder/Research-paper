from mas_reliab.utils import exact_equal, stable_seed


def test_exact_state_equality_is_key_order_invariant():
    assert exact_equal({"b": 2, "a": 1}, {"a": 1, "b": 2})
    assert not exact_equal({"a": 1}, {"a": 2})


def test_stable_seed():
    assert stable_seed("episode", 7) == stable_seed("episode", 7)
    assert stable_seed("episode", 7) != stable_seed("episode", 8)
