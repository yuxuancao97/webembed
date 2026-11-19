#!/bin/zsh
# Installs the WebEmbed manifest that loads from GitHub Pages.
cd "$(dirname "$0")"
ADDIN_URL=https://yuxuancao97.github.io/webembed node scripts/sideload.mjs
echo
echo "Now quit PowerPoint with Cmd+Q and reopen your deck."
read -s -k 1 "?Press any key to close."
