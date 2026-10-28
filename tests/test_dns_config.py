"""Resolver configuration regressions; never contact an external DNS server."""

from __future__ import annotations

import importlib.util
import ipaddress
import os
from pathlib import Path
import sys
import threading

import dns.resolver
import pytest

DEFAULT_MODULE = Path(__file__).with_name("engine_section.py")
if not DEFAULT_MODULE.exists():
    DEFAULT_MODULE = Path(__file__).resolve().parents[1] / "app.py"
MODULE_PATH = Path(os.environ.get("VERIFIER_MODULE", DEFAULT_MODULE))
spec = importlib.util.spec_from_file_location("verifier_dns_config_test", MODULE_PATH)
engine = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = engine
spec.loader.exec_module(engine)


class ConfigResolver:
    def __init__(self, servers):
        self._servers = list(servers)
        self.assignments = []

    @property
    def nameservers(self):
        return list(self._servers)

    @nameservers.setter
    def nameservers(self, servers):
        for server in servers:
            ipaddress.ip_address(server)
        self._servers = list(servers)
        self.assignments.append(list(servers))


@pytest.fixture(autouse=True)
def isolated_dns_configuration(monkeypatch):
    monkeypatch.delenv("VERIFIER_DNS_NAMESERVERS", raising=False)
    monkeypatch.setattr(engine, "_DNS_LOCAL", threading.local())


def test_missing_system_dns_uses_public_fallback_only(monkeypatch):
    resolver = ConfigResolver([])
    monkeypatch.setattr(engine.dns.resolver, "Resolver", lambda **_kwargs: resolver)
    assert engine.validate_dns_configuration() == ("1.1.1.1", "8.8.8.8")
    assert resolver.assignments == [["1.1.1.1", "8.8.8.8"]]
    assert resolver.timeout == 0.65 and resolver.lifetime == 2.0
    assert isinstance(resolver.cache, dns.resolver.LRUCache)


def test_available_system_dns_is_preserved(monkeypatch):
    resolver = ConfigResolver(["192.168.1.1"])
    monkeypatch.setattr(engine.dns.resolver, "Resolver", lambda **_kwargs: resolver)
    assert engine.validate_dns_configuration() == ("192.168.1.1",)
    assert resolver.assignments == []


def test_explicit_nameserver_override_is_applied_and_validated(monkeypatch):
    resolver = ConfigResolver(["192.168.1.1"])
    monkeypatch.setenv("VERIFIER_DNS_NAMESERVERS", " 9.9.9.9 , 149.112.112.112 ")
    monkeypatch.setattr(engine.dns.resolver, "Resolver", lambda **_kwargs: resolver)
    assert engine.validate_dns_configuration() == ("9.9.9.9", "149.112.112.112")
    assert resolver.assignments == [["9.9.9.9", "149.112.112.112"]]


@pytest.mark.parametrize("override", ["not-a-dns-server", ", ,", "1.1.1.1,bad"])
def test_malformed_override_raises_descriptive_configuration_error(
    monkeypatch, override
):
    monkeypatch.setenv("VERIFIER_DNS_NAMESERVERS", override)
    monkeypatch.setattr(
        engine.dns.resolver, "Resolver", lambda **_kwargs: ConfigResolver([])
    )
    with pytest.raises(ValueError, match="VERIFIER_DNS_NAMESERVERS"):
        engine.validate_dns_configuration()


def test_missing_resolver_configuration_is_recovered(monkeypatch):
    calls = []

    def factory(configure=True):
        calls.append(configure)
        if configure:
            raise dns.resolver.NoResolverConfiguration()
        return ConfigResolver([])

    monkeypatch.setattr(engine.dns.resolver, "Resolver", factory)
    assert engine.validate_dns_configuration() == ("1.1.1.1", "8.8.8.8")
    assert calls == [True, False]


def test_thread_local_resolver_reuses_config_and_rebuilds_for_override(monkeypatch):
    built = []

    def factory(**_kwargs):
        resolver = ConfigResolver(["192.168.1.1"])
        built.append(resolver)
        return resolver

    monkeypatch.setattr(engine.dns.resolver, "Resolver", factory)
    first = engine._thread_resolver()
    assert engine._thread_resolver() is first
    assert len(built) == 1
    monkeypatch.setenv("VERIFIER_DNS_NAMESERVERS", "9.9.9.9")
    second = engine._thread_resolver()
    assert second is not first
    assert second.nameservers == ["9.9.9.9"]
    assert len(built) == 2


def test_malformed_override_fails_before_database_or_lookup(monkeypatch, tmp_path):
    source = tmp_path / "leads.csv"
    source.write_text("Email\nalice@example.com\n", encoding="utf-8")
    monkeypatch.setenv("VERIFIER_DNS_NAMESERVERS", "bad-configuration")
    monkeypatch.setattr(
        engine.dns.resolver, "Resolver", lambda **_kwargs: ConfigResolver([])
    )
    with pytest.raises(ValueError, match="VERIFIER_DNS_NAMESERVERS"):
        engine.run_verification(source, tmp_path / "run", 0)
    assert not (tmp_path / "run" / "results.sqlite3").exists()
