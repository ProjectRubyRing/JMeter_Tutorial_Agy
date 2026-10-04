#!/bin/bash
# ==============================================================================
# JMeter Graceful Shutdown Script
# Sends a polite shutdown signal to the running JMeter instance, ensuring
# current requests complete and JTL result files are properly flushed and closed.
# ==============================================================================
set -euo pipefail

JMETER_HOME="/opt/jmeter/apache-jmeter-5.6.3"
SHUTDOWN_BIN="${JMETER_HOME}/bin/shutdown.sh"
STOPTEST_BIN="${JMETER_HOME}/bin/stoptest.sh"

MODE="${1:-graceful}"

if [ "${MODE}" = "immediate" ]; then
    echo "[INFO] Sending IMMEDIATE stop signal to JMeter..."
    if [ -f "${STOPTEST_BIN}" ]; then
        bash "${STOPTEST_BIN}"
    else
        echo "[ERROR] ${STOPTEST_BIN} not found." >&2
        exit 1
    fi
else
    echo "[INFO] Sending GRACEFUL shutdown signal to JMeter..."
    echo "[INFO] Active requests will be allowed to complete. Please wait..."
    if [ -f "${SHUTDOWN_BIN}" ]; then
        bash "${SHUTDOWN_BIN}"
    else
        echo "[ERROR] ${SHUTDOWN_BIN} not found." >&2
        exit 1
    fi
fi

echo "[INFO] Signal sent. Monitor JMeter terminal or log file to confirm shutdown."
