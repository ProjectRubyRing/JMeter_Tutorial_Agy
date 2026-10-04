#!/bin/bash
# ==============================================================================
# JMeter Non-GUI Load Test Automation Script
# Usage: ./run_load_test.sh [THREADS] [RAMPUP_SEC] [DURATION_SEC]
# ==============================================================================
set -euo pipefail

# 1. Configuration & Directories
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
BASE_DIR="/data/jmeter"
JMETER_BIN="/opt/jmeter/apache-jmeter-5.6.3/bin/jmeter"

SCENARIO_DIR="${BASE_DIR}/scenarios"
RESULT_DIR="${BASE_DIR}/results"
LOG_DIR="${BASE_DIR}/logs"

# Ensure directories exist
mkdir -p "${RESULT_DIR}" "${LOG_DIR}"

# 2. Parameters (Defaults if not supplied)
THREADS="${1:-50}"
RAMPUP="${2:-60}"
DURATION="${3:-1200}"

TIMESTAMP="$(date +'%Y%m%d_%H%M%S')"
SCENARIO_FILE="${SCENARIO_DIR}/scenario_ec_purchase.jmx"
RESULT_JTL="${RESULT_DIR}/run_${TIMESTAMP}.jtl"
JMETER_LOG="${LOG_DIR}/jmeter_${TIMESTAMP}.log"
HTML_REPORT_DIR="${RESULT_DIR}/html_${TIMESTAMP}"

# 3. Validation
if [ ! -f "${JMETER_BIN}" ]; then
    echo "[ERROR] JMeter executable not found at ${JMETER_BIN}" >&2
    exit 1
fi

if [ ! -f "${SCENARIO_FILE}" ]; then
    echo "[WARNING] Scenario file ${SCENARIO_FILE} not found."
    echo "Searching for any .jmx in ${SCENARIO_DIR}..."
    FIRST_JMX=$(find "${SCENARIO_DIR}" -maxdepth 1 -name "*.jmx" | head -n 1)
    if [ -n "${FIRST_JMX}" ]; then
        SCENARIO_FILE="${FIRST_JMX}"
        echo "[INFO] Using detected scenario: ${SCENARIO_FILE}"
    else
        echo "[ERROR] No .jmx scenario found in ${SCENARIO_DIR}." >&2
        exit 1
    fi
fi

echo "=========================================================="
echo " Starting JMeter Non-GUI Performance Test"
echo " Timestamp      : ${TIMESTAMP}"
echo " Scenario       : ${SCENARIO_FILE}"
echo " Threads (Users): ${THREADS}"
echo " Ramp-up (s)    : ${RAMPUP}"
echo " Duration (s)   : ${DURATION}"
echo " JTL Output     : ${RESULT_JTL}"
echo " JMeter Log     : ${JMETER_LOG}"
echo " HTML Report    : ${HTML_REPORT_DIR}"
echo "=========================================================="

# 4. Execute JMeter CLI
"${JMETER_BIN}" -n \
    -t "${SCENARIO_FILE}" \
    -l "${RESULT_JTL}" \
    -j "${JMETER_LOG}" \
    -e -o "${HTML_REPORT_DIR}" \
    -Jthreads="${THREADS}" \
    -Jrampup="${RAMPUP}" \
    -Jduration="${DURATION}"

EXIT_CODE=$?

echo "=========================================================="
if [ ${EXIT_CODE} -eq 0 ]; then
    echo " [SUCCESS] JMeter load test completed successfully."
    echo " Results saved to:"
    echo "   JTL : ${RESULT_JTL}"
    echo "   Log : ${JMETER_LOG}"
    echo "   HTML: ${HTML_REPORT_DIR}/index.html"
else
    echo " [FAILURE] JMeter exited with status code ${EXIT_CODE}."
    echo " Check log for details: ${JMETER_LOG}"
fi
echo "=========================================================="

exit ${EXIT_CODE}
