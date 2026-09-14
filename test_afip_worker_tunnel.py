# -*- coding: utf-8 -*-
"""Unit tests for named vs quick Cloudflare tunnel (no live Cloudflare account)."""
from __future__ import annotations

import io
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from afip_worker import tunnel


ENV_KEYS = (
    "AFIP_CLOUDFLARED_CONFIG",
    "AFIP_TUNNEL_HOSTNAME",
    "AFIP_WORKER_PUBLIC_URL",
    "AFIP_CLOUDFLARED_TOKEN",
    "TUNNEL_TOKEN",
    "CLOUDFLARED",
)

SAMPLE_CONFIG = """
tunnel: 11111111-2222-3333-4444-555555555555
credentials-file: jobs/cloudflared/11111111-2222-3333-4444-555555555555.json

ingress:
  - hostname: afip-worker.example.com  # público
    service: http://127.0.0.1:8765
  - service: http_status:404
"""


class _EnvIsolated(unittest.TestCase):
    def setUp(self) -> None:
        self._saved = {k: os.environ.get(k) for k in ENV_KEYS}
        for k in ENV_KEYS:
            os.environ.pop(k, None)

    def tearDown(self) -> None:
        for k in ENV_KEYS:
            os.environ.pop(k, None)
        for k, v in self._saved.items():
            if v is None:
                os.environ.pop(k, None)
            else:
                os.environ[k] = v


class ParseConfigTests(_EnvIsolated):
    def test_parse_ingress_hostname(self) -> None:
        host = tunnel.parse_ingress_hostname(SAMPLE_CONFIG)
        self.assertEqual(host, "afip-worker.example.com")

    def test_parse_quoted_hostname(self) -> None:
        text = 'ingress:\n  - hostname: "arca.estudio.com"\n    service: http://127.0.0.1:8765\n'
        self.assertEqual(tunnel.parse_ingress_hostname(text), "arca.estudio.com")

    def test_skips_localhost(self) -> None:
        text = "hostname: localhost\nhostname: real.example.com\n"
        self.assertEqual(tunnel.parse_ingress_hostname(text), "real.example.com")

    def test_parse_local_service(self) -> None:
        self.assertEqual(tunnel.parse_local_service(SAMPLE_CONFIG), "http://127.0.0.1:8765")

    def test_normalize_public_url(self) -> None:
        self.assertEqual(
            tunnel.normalize_public_url("afip-worker.example.com"),
            "https://afip-worker.example.com",
        )
        self.assertEqual(
            tunnel.normalize_public_url("https://afip-worker.example.com/"),
            "https://afip-worker.example.com",
        )
        self.assertEqual(tunnel.normalize_public_url("  "), "")


class DetectNamedTunnelTests(_EnvIsolated):
    def test_example_template_hostname(self) -> None:
        example = Path(__file__).resolve().parent / "jobs" / "cloudflared" / "config.yml.example"
        text = example.read_text(encoding="utf-8")
        self.assertEqual(tunnel.parse_ingress_hostname(text), "afip-worker.TU-DOMINIO.com")
        self.assertEqual(tunnel.parse_local_service(text), "http://127.0.0.1:8765")
        self.assertIsNone(tunnel.detect_named_tunnel())

    def test_absent_config_is_quick_tunnel(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            with patch.object(tunnel, "_repo_root", return_value=root):
                self.assertIsNone(tunnel.detect_named_tunnel())

    def test_default_jobs_cloudflared_config(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            cfg_dir = root / "jobs" / "cloudflared"
            cfg_dir.mkdir(parents=True)
            (cfg_dir / "config.yml").write_text(SAMPLE_CONFIG, encoding="utf-8")
            uuid_json = cfg_dir / "11111111-2222-3333-4444-555555555555.json"
            uuid_json.write_text("{}", encoding="utf-8")
            with patch.object(tunnel, "_repo_root", return_value=root):
                spec = tunnel.detect_named_tunnel()
        self.assertIsNotNone(spec)
        assert spec is not None
        self.assertEqual(spec.kind, "config")
        self.assertEqual(spec.public_url, "https://afip-worker.example.com")
        self.assertIsNone(spec.error)
        self.assertTrue(str(spec.config_path).endswith("config.yml"))

    def test_env_config_path_preferred(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            other = root / "custom.yml"
            other.write_text(
                "ingress:\n  - hostname: custom.example.com\n    service: http://127.0.0.1:8765\n",
                encoding="utf-8",
            )
            default_dir = root / "jobs" / "cloudflared"
            default_dir.mkdir(parents=True)
            (default_dir / "config.yml").write_text(SAMPLE_CONFIG, encoding="utf-8")
            os.environ["AFIP_CLOUDFLARED_CONFIG"] = str(other)
            with patch.object(tunnel, "_repo_root", return_value=root):
                spec = tunnel.detect_named_tunnel()
        assert spec is not None
        self.assertEqual(spec.public_url, "https://custom.example.com")
        self.assertEqual(spec.config_path, other)

    def test_hostname_override(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            cfg = Path(raw) / "config.yml"
            cfg.write_text(SAMPLE_CONFIG, encoding="utf-8")
            os.environ["AFIP_CLOUDFLARED_CONFIG"] = str(cfg)
            os.environ["AFIP_TUNNEL_HOSTNAME"] = "https://override.example.com"
            spec = tunnel.detect_named_tunnel()
        assert spec is not None
        self.assertEqual(spec.public_url, "https://override.example.com")

    def test_missing_env_config_is_error_not_quick(self) -> None:
        os.environ["AFIP_CLOUDFLARED_CONFIG"] = "/no/such/cloudflared-config.yml"
        spec = tunnel.detect_named_tunnel()
        assert spec is not None
        self.assertIsNotNone(spec.error)
        self.assertIn("no existe", spec.error or "")

    def test_token_requires_hostname(self) -> None:
        os.environ["AFIP_CLOUDFLARED_TOKEN"] = "secret-token"
        spec = tunnel.detect_named_tunnel()
        assert spec is not None
        self.assertEqual(spec.kind, "token")
        self.assertIsNotNone(spec.error)
        os.environ["AFIP_TUNNEL_HOSTNAME"] = "zt.example.com"
        spec = tunnel.detect_named_tunnel()
        assert spec is not None
        self.assertEqual(spec.public_url, "https://zt.example.com")
        self.assertIsNone(spec.error)


class CommandTests(_EnvIsolated):
    def test_named_config_command(self) -> None:
        spec = tunnel.NamedTunnelSpec(
            kind="config",
            public_url="https://afip-worker.example.com",
            config_path=Path("/tmp/config.yml"),
        )
        cmd = tunnel.named_tunnel_command("cloudflared", spec)
        self.assertEqual(
            cmd,
            ["cloudflared", "tunnel", "--no-autoupdate", "--config", "/tmp/config.yml", "run"],
        )
        self.assertNotIn("--url", cmd)

    def test_named_token_command_redacted_in_log(self) -> None:
        spec = tunnel.NamedTunnelSpec(
            kind="token",
            public_url="https://zt.example.com",
            token="super-secret",
        )
        cmd = tunnel.named_tunnel_command("cloudflared", spec)
        self.assertEqual(
            cmd,
            ["cloudflared", "tunnel", "--no-autoupdate", "run", "--token", "super-secret"],
        )
        logged = tunnel.cmd_for_log(cmd)
        self.assertIn("--token <redacted>", logged)
        self.assertNotIn("super-secret", logged)

    def test_quick_command_targets_local_port(self) -> None:
        cmd = tunnel.quick_tunnel_command("cloudflared", 9000)
        self.assertEqual(
            cmd,
            ["cloudflared", "tunnel", "--no-autoupdate", "--url", "http://127.0.0.1:9000"],
        )


class DummyProc:
    def __init__(self, lines: list[str] | None = None, running: bool = True) -> None:
        text = "\n".join(lines or ["Registered tunnel connection"]) + "\n"
        self.stdout = io.StringIO(text)
        self._rc = None if running else 1

    def poll(self) -> int | None:
        return self._rc


class StartTunnelTests(_EnvIsolated):
    def test_named_start_writes_stable_url_not_trycloudflare(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            cfg = root / "config.yml"
            cfg.write_text(
                "ingress:\n"
                "  - hostname: afip-worker.example.com\n"
                "    service: http://127.0.0.1:8765\n"
                "  - service: http_status:404\n",
                encoding="utf-8",
            )
            bridge = root / "cloud_bridge.txt"
            os.environ["AFIP_CLOUDFLARED_CONFIG"] = str(cfg)
            spawned: list[list[str]] = []

            def fake_spawn(cmd: list[str]) -> DummyProc:
                spawned.append(cmd)
                return DummyProc()

            with (
                patch.object(tunnel, "find_cloudflared", return_value="cloudflared"),
                patch.object(tunnel, "_spawn", side_effect=fake_spawn),
                patch.object(tunnel, "_drain_stdout", lambda proc: None),
                patch.object(tunnel, "bridge_info_path", return_value=bridge),
            ):
                proc = tunnel.start_cloudflared_tunnel(
                    port=8765,
                    token="tok",
                    named_startup_wait=0,
                )
            self.assertIsNotNone(proc)
            self.assertEqual(len(spawned), 1)
            self.assertIn("--config", spawned[0])
            self.assertNotIn("--url", spawned[0])
            text = bridge.read_text(encoding="utf-8")
            self.assertIn('AFIP_WORKER_URL = "https://afip-worker.example.com"', text)
            self.assertIn('AFIP_WORKER_TOKEN = "tok"', text)
            self.assertIn("URL estable", text)
            self.assertNotIn("trycloudflare", text)

    def test_quick_start_parses_trycloudflare_url(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            bridge = Path(raw) / "cloud_bridge.txt"
            spawned: list[list[str]] = []

            def fake_spawn(cmd: list[str]) -> DummyProc:
                spawned.append(cmd)
                return DummyProc(
                    lines=[
                        "Thank you for trying Cloudflare Tunnel",
                        "https://random-name.trycloudflare.com",
                    ]
                )

            with (
                patch.object(tunnel, "find_cloudflared", return_value="cloudflared"),
                patch.object(tunnel, "detect_named_tunnel", return_value=None),
                patch.object(tunnel, "_spawn", side_effect=fake_spawn),
                patch.object(tunnel, "_drain_stdout", lambda proc: None),
                patch.object(tunnel, "bridge_info_path", return_value=bridge),
            ):
                proc = tunnel.start_cloudflared_tunnel(
                    port=8765,
                    token="tok",
                    quick_url_wait=2,
                )
            self.assertIsNotNone(proc)
            self.assertIn("--url", spawned[0])
            text = bridge.read_text(encoding="utf-8")
            self.assertIn("https://random-name.trycloudflare.com", text)
            self.assertIn("túnel rápido cambia de URL", text)

    def test_write_bridge_info_keeps_token(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            bridge = Path(raw) / "cloud_bridge.txt"
            with patch.object(tunnel, "bridge_info_path", return_value=bridge):
                tunnel.write_bridge_info(
                    url="https://afip-worker.example.com",
                    token="keep-me",
                    port=8765,
                    stable=True,
                )
            text = bridge.read_text(encoding="utf-8")
            self.assertIn("keep-me", text)
            self.assertIn("8765", text)


if __name__ == "__main__":
    unittest.main()
