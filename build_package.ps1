if (-not (Test-Path "office-vba-reference\.git")) {
    git submodule update --init
}

python -m scripts
python -m pip install -e .
python -m tests
python -m build
