#!/bin/bash

src_dir="./src"
output_file="concat-script.js"

# Delete the output file if it exists
if [ -f "$output_file" ]; then
  rm "$output_file"
fi

touch "$output_file"

append_with_newline() {
  # Append file content excluding lines with 'module.exports'
  sed '/module\.exports/d' "$1" >> "$output_file"
  echo -e "\n" >> "$output_file"  # Adds an extra newline after each file
}

echo "// Controllers" >> "$output_file"
append_with_newline "$src_dir/controllers/SheetController.js"
append_with_newline "$src_dir/controllers/EventController.js"

echo "// Services" >> "$output_file"
append_with_newline "$src_dir/services/SheetService.js"
append_with_newline "$src_dir/services/DropdownService.js"
append_with_newline "$src_dir/services/ProtectionService.js"
append_with_newline "$src_dir/services/WordCountService.js"
append_with_newline "$src_dir/services/LanguageService.js"
append_with_newline "$src_dir/services/MenuService.js"

echo "// Utils" >> "$output_file"
append_with_newline "$src_dir/utils/Utils.js"

echo "// Main Script" >> "$output_file"
append_with_newline "$src_dir/Main.js"

echo "Concatenation complete. Output written to $output_file"
