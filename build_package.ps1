if (-not (Test-Path "office-vba-reference\.git")) {
    git submodule update --init
}

python -m scripts
python -m pip install -e .
# TODO run tests
python -m build

