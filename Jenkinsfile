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
              // Create both the main and .streamlit directories first
              sh "ssh -o StrictHostKeyChecking=no ubuntu@${EC2_IP_ADDRESS} 'mkdir -p /home/ubuntu/BCA_Streamlit/.streamlit/'"
              // Now copy the secrets file
              sh """
                scp -o StrictHostKeyChecking=no \$SECRETS_FILE ubuntu@${EC2_IP_ADDRESS}:/home/ubuntu/BCA_Streamlit/.streamlit/secrets.toml
              """
              // Copy the rest of the project
              sh "scp -o StrictHostKeyChecking=no -r ./* ubuntu@${EC2_IP_ADDRESS}:/home/ubuntu/BCA_Streamlit/"
              // Run the preparation script
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
                cd /home/ubuntu/BCA_Streamlit &&
                source venv/bin/activate &&
                pkill -f "streamlit run" || true &&
                nohup venv/bin/streamlit run main.py --server.port=8501 --server.headless true > streamlit.log 2>&1 & &&
                ps aux | grep streamlit
              '
            """
          }
        }
      }
    }
  }
}