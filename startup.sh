echo "Installing Homebrew..."
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
echo "Updating Homebrew..."
brew update && brew upgrade
echo "Installing Python 3..."
brew install python3
echo "Installing script dependancies..."
brew install pyqt@5
pip install docx2txt
pip install python-docx