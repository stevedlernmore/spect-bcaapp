pipeline {
  agent any

  environment {
    EC2_IP_ADDRESS = '13.213.50.58'
  }

  stages {
    stage('Prepare EC2 Instance') {
      steps {
        script {
          echo 'Preparing EC2 instance...'
          def prepareScript = 'prepare_ec2.sh'
          def prepareExecute = "bash ./${prepareScript}"
          withCredentials([file(credentialsId: 'BCA-secrets', variable: 'SECRETS_FILE')]) {
            sshagent(['spectra-ec2']) {
              sh """
                ssh -o StrictHostKeyChecking=no ubuntu@${EC2_IP_ADDRESS} '
                  sudo mkdir -p /home/ubuntu/BCA_Streamlit/.streamlit &&
                  sudo chown -R ubuntu:ubuntu /home/ubuntu/BCA_Streamlit &&
                  sudo chmod -R u+rwX /home/ubuntu/BCA_Streamlit
                '
              """
              sh """
                scp -o StrictHostKeyChecking=no \$SECRETS_FILE ubuntu@${EC2_IP_ADDRESS}:/home/ubuntu/BCA_Streamlit/.streamlit/secrets.toml
              """
              sh "scp -o StrictHostKeyChecking=no -r . ubuntu@${EC2_IP_ADDRESS}:/home/ubuntu/BCA_Streamlit/"
              sh "ssh -o StrictHostKeyChecking=no ubuntu@${EC2_IP_ADDRESS} 'cd /home/ubuntu/BCA_Streamlit && bash ${prepareScript}'"
            }
          }
        }
      }
    }
    stage('Deploy Streamlit App') {
      steps {
        script {
          echo 'Deploying Streamlit app...'
          sshagent(['spectra-ec2']) {
            sh """
              ssh -o StrictHostKeyChecking=no ubuntu@${EC2_IP_ADDRESS} '
                sudo systemctl daemon-reload &&
                sudo systemctl restart bca_streamlit &&
                sudo systemctl enable bca_streamlit
              '
            """
          }
        }
      }
    }
    stage('Configure Nginx for HTTPS') {
      steps {
        script {
          echo 'Configuring Nginx as HTTPS reverse proxy...'
          sshagent(['spectra-ec2']) {
            sh """
ssh -o StrictHostKeyChecking=no ubuntu@${EC2_IP_ADDRESS} 'bash -s' << 'EOF'
sudo tee /etc/nginx/sites-available/streamlit > /dev/null << 'NGINX_CONF'
server {
  listen 80;
  server_name bca-trial.xyz 13.213.50.58;

  location / {
    proxy_pass http://localhost:8501;
    proxy_set_header Host \$host;
    proxy_set_header X-Real-IP \$remote_addr;
    proxy_set_header X-Forwarded-For \$proxy_add_x_forwarded_for;
    proxy_set_header X-Forwarded-Proto \$scheme;
    proxy_http_version 1.1;
    proxy_set_header Upgrade \$http_upgrade;
    proxy_set_header Connection "upgrade";
  }
}
NGINX_CONF
EOF
"""
            sh """
              ssh -o StrictHostKeyChecking=no ubuntu@${EC2_IP_ADDRESS} '
                sudo ln -sf /etc/nginx/sites-available/streamlit /etc/nginx/sites-enabled/streamlit &&
                sudo rm -f /etc/nginx/sites-enabled/default &&
                sudo nginx -t &&
                sudo systemctl reload nginx
              '
            """

            sh """
              ssh -o StrictHostKeyChecking=no ubuntu@${EC2_IP_ADDRESS} '
                sudo certbot --nginx --non-interactive --agree-tos -m aaronjohn.tamayo29@gmail.com -d bca-trial.xyz
              '
            """
          }
        }
      }
    }
  }
}