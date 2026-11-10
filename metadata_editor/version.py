"""The release workflow writes _release_version.py from the published tag."""

try:
    from ._release_version import APP_VERSION
except ModuleNotFoundError:
    APP_VERSION = "4.2.0-dev"
