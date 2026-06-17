"""Tests for the Saved Variance Profiles data layer."""
from __future__ import annotations

import pytest

from src.variance_profiles import (
    ACTION_CREATE,
    ACTION_DELETE,
    ACTION_DUPLICATE,
    ACTION_RESET_DEFAULT,
    ACTION_SET_DEFAULT,
    ACTION_UPDATE,
    PROFILE_TYPE_DELIVERY,
    PROFILE_TYPE_SALES_ORDER,
    Context,
    ProfileNotFoundError,
    ProfileStore,
    ProfileValidationError,
    SystemProfileError,
    validate_filters,
)


@pytest.fixture
def store(tmp_path):
    db = tmp_path / "profiles.db"
    return ProfileStore(db_path=str(db), ctx=Context(tenant_id="t1", user_id="alice"))


# -- system profiles ---------------------------------------------------------
def test_system_profiles_seeded(store):
    deliv = store.list_profiles(PROFILE_TYPE_DELIVERY)
    so = store.list_profiles(PROFILE_TYPE_SALES_ORDER)
    assert len(deliv) == 5 and all(p.is_system for p in deliv)
    assert len(so) == 5 and all(p.is_system for p in so)
    assert "All Delivery Variances" in {p.name for p in deliv}
    assert "All Sales Order Variances" in {p.name for p in so}


def test_seeding_is_idempotent(tmp_path):
    db = tmp_path / "p.db"
    ctx = Context(tenant_id="t1", user_id="alice")
    ProfileStore(db_path=str(db), ctx=ctx)
    s2 = ProfileStore(db_path=str(db), ctx=ctx)
    assert len(s2.list_profiles(PROFILE_TYPE_DELIVERY)) == 5


def test_system_profiles_cannot_be_edited_or_deleted(store):
    sysp = store.list_profiles(PROFILE_TYPE_DELIVERY)[0]
    with pytest.raises(SystemProfileError):
        store.update_profile(sysp.id, filters={})
    with pytest.raises(SystemProfileError):
        store.delete_profile(sysp.id)


def test_system_profiles_can_be_duplicated(store):
    sysp = next(p for p in store.list_profiles(PROFILE_TYPE_DELIVERY) if p.is_system)
    dup = store.duplicate_profile(sysp.id, "My Copy")
    assert dup.is_system is False
    assert dup.name == "My Copy"


# -- CRUD --------------------------------------------------------------------
def test_create_and_get(store):
    p = store.create_profile("My Delivery Variance", PROFILE_TYPE_DELIVERY,
                             {"warehouses": ["2910"], "variance_amount_threshold": 1000})
    fetched = store.get_profile(p.id)
    assert fetched.name == "My Delivery Variance"
    assert fetched.filters["warehouses"] == ["2910"]
    assert fetched.is_system is False


def test_duplicate_name_rejected(store):
    store.create_profile("Dupe", PROFILE_TYPE_DELIVERY, {})
    with pytest.raises(ProfileValidationError):
        store.create_profile("Dupe", PROFILE_TYPE_DELIVERY, {})


def test_update(store):
    p = store.create_profile("P", PROFILE_TYPE_DELIVERY, {"warehouses": ["2910"]})
    updated = store.update_profile(p.id, filters={"warehouses": ["2920", "2930"]}, name="P2")
    assert updated.name == "P2"
    assert updated.filters["warehouses"] == ["2920", "2930"]


def test_delete(store):
    p = store.create_profile("Gone", PROFILE_TYPE_DELIVERY, {})
    store.delete_profile(p.id)
    with pytest.raises(ProfileNotFoundError):
        store.get_profile(p.id)


def test_duplicate_auto_names(store):
    p = store.create_profile("Base", PROFILE_TYPE_DELIVERY, {})
    d1 = store.duplicate_profile(p.id)
    d2 = store.duplicate_profile(p.id)
    assert d1.name == "Base (copy)"
    assert d2.name == "Base (copy) 2"


# -- defaults ----------------------------------------------------------------
def test_default_falls_back_to_system(store):
    d = store.get_default_profile(PROFILE_TYPE_DELIVERY)
    assert d.name == "All Delivery Variances" and d.is_system


def test_set_default_is_exclusive(store):
    a = store.create_profile("A", PROFILE_TYPE_DELIVERY, {})
    b = store.create_profile("B", PROFILE_TYPE_DELIVERY, {})
    store.set_default(a.id)
    store.set_default(b.id)
    defaults = [p for p in store.list_profiles(PROFILE_TYPE_DELIVERY) if p.is_default]
    assert len(defaults) == 1 and defaults[0].id == b.id
    assert store.get_default_profile(PROFILE_TYPE_DELIVERY).id == b.id


def test_reset_default(store):
    a = store.create_profile("A", PROFILE_TYPE_DELIVERY, {})
    store.set_default(a.id)
    back = store.reset_default(PROFILE_TYPE_DELIVERY)
    assert back.name == "All Delivery Variances"
    assert not any(p.is_default for p in store.list_profiles(PROFILE_TYPE_DELIVERY))


def test_default_is_per_type(store):
    a = store.create_profile("A", PROFILE_TYPE_DELIVERY, {})
    b = store.create_profile("B", PROFILE_TYPE_SALES_ORDER, {})
    store.set_default(a.id)
    store.set_default(b.id)
    assert store.get_default_profile(PROFILE_TYPE_DELIVERY).id == a.id
    assert store.get_default_profile(PROFILE_TYPE_SALES_ORDER).id == b.id


# -- isolation ---------------------------------------------------------------
def test_user_isolation(tmp_path):
    db = tmp_path / "p.db"
    alice = ProfileStore(db_path=str(db), ctx=Context(tenant_id="t1", user_id="alice"))
    bob = ProfileStore(db_path=str(db), ctx=Context(tenant_id="t1", user_id="bob"))
    alice.create_profile("Alice Only", PROFILE_TYPE_DELIVERY, {})
    bob_user = [p for p in bob.list_profiles(PROFILE_TYPE_DELIVERY) if not p.is_system]
    assert bob_user == []


def test_tenant_isolation(tmp_path):
    db = tmp_path / "p.db"
    t1 = ProfileStore(db_path=str(db), ctx=Context(tenant_id="t1", user_id="alice"))
    t2 = ProfileStore(db_path=str(db), ctx=Context(tenant_id="t2", user_id="alice"))
    t1.create_profile("Tenant1 Only", PROFILE_TYPE_DELIVERY, {})
    t2_user = [p for p in t2.list_profiles(PROFILE_TYPE_DELIVERY) if not p.is_system]
    assert t2_user == []


# -- validation --------------------------------------------------------------
def test_validate_filters_rejects_bad_types():
    with pytest.raises(ProfileValidationError):
        validate_filters({"warehouses": 5})
    with pytest.raises(ProfileValidationError):
        validate_filters({"variance_amount_threshold": "abc"})
    with pytest.raises(ProfileValidationError):
        validate_filters({"sort_direction": "sideways"})


def test_validate_filters_preserves_unknown_under_extra():
    out = validate_filters({"some_future_field": [1, 2]})
    assert out["extra"]["some_future_field"] == [1, 2]


def test_empty_name_rejected(store):
    with pytest.raises(ProfileValidationError):
        store.create_profile("   ", PROFILE_TYPE_DELIVERY, {})


# -- audit -------------------------------------------------------------------
def test_audit_log_records_all_actions(store):
    p = store.create_profile("Aud", PROFILE_TYPE_DELIVERY, {})
    store.update_profile(p.id, filters={"warehouses": ["2910"]})
    d = store.duplicate_profile(p.id)
    store.set_default(p.id)
    store.reset_default(PROFILE_TYPE_DELIVERY)
    store.delete_profile(d.id)
    actions = {a["action"] for a in store.list_audit()}
    assert {ACTION_CREATE, ACTION_UPDATE, ACTION_DUPLICATE,
            ACTION_SET_DEFAULT, ACTION_RESET_DEFAULT, ACTION_DELETE} <= actions
