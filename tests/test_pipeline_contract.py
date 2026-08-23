import subprocess
import sys

import run_pipeline


def test_stage_runs_from_repository_root(monkeypatch, tmp_path):
    calls = []

    def fake_run(command, *, cwd, check):
        calls.append((command, cwd, check))

    monkeypatch.setattr(run_pipeline.subprocess, "run", fake_run)
    script = tmp_path / "stage.py"

    assert run_pipeline.run_command(script, "test stage", project_root=tmp_path)
    assert calls == [([sys.executable, str(script)], tmp_path, True)]


def test_stage_failure_is_reported_without_running_following_stages(
    monkeypatch, tmp_path
):
    error = subprocess.CalledProcessError(returncode=7, cmd=[sys.executable])

    def fake_run(*args, **kwargs):
        raise error

    monkeypatch.setattr(run_pipeline.subprocess, "run", fake_run)

    assert not run_pipeline.run_command(
        tmp_path / "stage.py", "test stage", project_root=tmp_path
    )


def test_main_resolves_all_stages_relative_to_the_script(monkeypatch, tmp_path):
    for script_name, _ in run_pipeline.STAGES:
        (tmp_path / script_name).touch()

    monkeypatch.setattr(run_pipeline, "PROJECT_ROOT", tmp_path)
    calls = []

    def fake_run_command(script_path, description):
        calls.append((script_path, description))
        return True

    monkeypatch.setattr(run_pipeline, "run_command", fake_run_command)

    assert run_pipeline.main() == 0
    assert calls == [
        (tmp_path / script_name, description)
        for script_name, description in run_pipeline.STAGES
    ]
