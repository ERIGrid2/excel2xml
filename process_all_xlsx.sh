#!/bin/sh

set -o errexit

if [ -z "$1" ]
then
  OUTPUT_DIR="/github/workspace/erigrid2-test-cases"
else
  OUTPUT_DIR="$1"
fi

mkdir -p ${OUTPUT_DIR}

RUN_EXCEL2MD=$(cat << EOM
  FILE=\$0
  OUTPUT_DIR="\$1"
  PREFIX="\$2"
  DIRNAMEPREFIX=\$(dirname "\${FILE}")
  IS_TLD=\$(echo \${DIRNAMEPREFIX} | grep "/")
  if [ \$? -eq 1 ]; then
    DIRNAME=/
  else
    DIRNAME=/\${DIRNAMEPREFIX#\$PREFIX}
  fi
  mkdir -p "\${OUTPUT_DIR}/\${DIRNAME}/"
  OUTPUT_FILE_NAME="\${OUTPUT_DIR}\${DIRNAME}/index.md"
  echo "Creating markdown file: \$OUTPUT_FILE_NAME"
  python3 xlsx2md.py "\$FILE" > "\${OUTPUT_FILE_NAME}"
  IMAGES=\$(find "\${DIRNAMEPREFIX}" -iname *.png)
  for IMAGE in \$IMAGES
  do
    echo Copying image file "\$IMAGE" into "\${OUTPUT_DIR}/\${DIRNAME}"
    cp "\$IMAGE" "\${OUTPUT_DIR}/\${DIRNAME}"
  done
EOM
)

# process all *.xlsx files from grupoetra and create index.md
find excel-input/* -type f -name '*.xlsx' -exec sh -c "$RUN_EXCEL2MD" {} ${OUTPUT_DIR} "excel-input/" ';'

find "${OUTPUT_DIR}" -type d -exec sh -c '
  DIRPATH=$0
  DIRNAME=$(basename "$DIRPATH")
  if [ ! -f "${DIRPATH}/index.md" ]
  then
    if [ ! -f "${DIRPATH}/_index.md" ]
    then
      echo Creating title link for directory: $DIRPATH with title: $DIRNAME
      cat > "${DIRPATH}/_index.md" <<EOF
---
title: "$DIRNAME"
linkTitle: "$DIRNAME"
weight: 5
---
EOF
    fi
  fi
' {} ${OUTPUT_DIR} ';'
