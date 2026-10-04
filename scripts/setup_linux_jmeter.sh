#!/bin/bash
# ==============================================================================
# JMeter Linux Server Environment Setup Script
# Installs Java 17, downloads JMeter 5.6.3, configures setenv.sh and OS limits.
# ==============================================================================
set -euo pipefail

JMETER_VERSION="5.6.3"
INSTALL_BASE="/opt/jmeter"
DATA_BASE="/data/jmeter"

echo "=== [1/5] Checking and Installing Java OpenJDK 17 ==="
if ! command -v java &> /dev/null; then
    if [ -f /etc/redhat-release ]; then
        echo "Detected RHEL/CentOS/Rocky. Installing java-17-openjdk..."
        sudo dnf install -y java-17-openjdk java-17-openjdk-devel wget tar
    elif [ -f /etc/debian_version ]; then
        echo "Detected Debian/Ubuntu. Installing openjdk-17-jdk..."
        sudo apt-get update && sudo apt-get install -y openjdk-17-jdk wget tar
    fi
else
    echo "Java is already installed: $(java -version 2>&1 | head -n 1)"
fi

echo "=== [2/5] Creating Directory Structure ==="
sudo mkdir -p "${INSTALL_BASE}"
sudo mkdir -p "${DATA_BASE}/scenarios" "${DATA_BASE}/csv" "${DATA_BASE}/results" "${DATA_BASE}/logs" "${DATA_BASE}/scripts"
sudo chown -R "${USER}:${USER}" "${INSTALL_BASE}" "${DATA_BASE}"

echo "=== [3/5] Downloading and Extracting Apache JMeter ${JMETER_VERSION} ==="
if [ ! -d "${INSTALL_BASE}/apache-jmeter-${JMETER_VERSION}" ]; then
    cd /tmp
    wget -q --show-progress "https://archive.apache.org/dist/jmeter/binaries/apache-jmeter-${JMETER_VERSION}.tgz"
    tar -zxvf "apache-jmeter-${JMETER_VERSION}.tgz" -C "${INSTALL_BASE}/"
    rm -f "apache-jmeter-${JMETER_VERSION}.tgz"
else
    echo "JMeter already exists at ${INSTALL_BASE}/apache-jmeter-${JMETER_VERSION}"
fi

echo "=== [4/5] Configuring JVM Settings (setenv.sh) ==="
SETENV_FILE="${INSTALL_BASE}/apache-jmeter-${JMETER_VERSION}/bin/setenv.sh"
cat << 'EOF' > "${SETENV_FILE}"
#!/bin/bash
# Optimize JVM for Load Testing
export JVM_ARGS="-Xms4g -Xmx4g -XX:MaxMetaspaceSize=512m -XX:+UseG1GC -XX:MaxGCPauseMillis=100 -XX:G1ReservePercent=15"
EOF
chmod +x "${SETENV_FILE}"
chmod +x "${INSTALL_BASE}/apache-jmeter-${JMETER_VERSION}/bin"/*

echo "=== [5/5] Configuring OS Limits (ulimit & sysctl) ==="
# Add limits if not already present
if ! grep -q "nofile 65535" /etc/security/limits.conf; then
    echo "* soft nofile 65535" | sudo tee -a /etc/security/limits.conf
    echo "* hard nofile 65535" | sudo tee -a /etc/security/limits.conf
fi

echo "=== Setup Complete! ==="
echo "Please add JMeter to your PATH by running:"
echo "  export PATH=\$PATH:${INSTALL_BASE}/apache-jmeter-${JMETER_VERSION}/bin"
echo "  echo 'export PATH=\$PATH:${INSTALL_BASE}/apache-jmeter-${JMETER_VERSION}/bin' >> ~/.bashrc"
