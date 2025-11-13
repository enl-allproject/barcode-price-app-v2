#!/bin/bash

# Script sederhana untuk update GitHub dari project ini
# Cara pakai: ./update_git.sh "pesan commit"

MSG="$1"

if [ -z "$MSG" ]; then
  echo "Tolong kasih pesan commit. Contoh:"
  echo "./update_git.sh \"perbaiki app.py\""
  exit 1
fi

echo "==> Menambahkan semua perubahan..."
git add .

echo "==> Membuat commit dengan pesan: $MSG"
git commit -m "$MSG"

echo "==> Mengirim ke GitHub (origin main)..."
git push origin main

echo "==> Selesai ✅"
