"""Smoke test: the UI module imports and exposes the panel entrypoint."""
def test_ui_module_imports():
    import src.column_variants_ui as ui
    assert callable(ui.render_variant_panel)
    assert callable(ui.get_store)
