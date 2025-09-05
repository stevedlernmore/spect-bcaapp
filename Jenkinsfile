pipeline {
  agent any

  environment {
    EC2_IP_ADDRESS = '13.213.50.58'
  }

  stages {
    stage('Fetch and Checkout Streamlit Repo') {
      steps {
        script{
          echo 'Fetching and checking out the Streamlit repository...'
          git branch: 'main', url: 'git@github.com:youruser/your-streamlit-repo.git'
        }
      }
    }
    stage('Prepare EC2 Instance') {
      steps {
        script {
          echo 'Preparing EC2 instance...'
          def prepareScript = 'prepare_ec2.sh'
          def prepareExecute = "bash ./${prepareScript}"
          sshagent(['spectra-ec2']) {
            sh "scp -o StrictHostKeyChecking=no ${prepareScript} ubuntu@${EC2_IP_ADDRESS}:/home/ubuntu/"
            sh "ssh -o StrictHostKeyChecking=no ubuntu@${EC2_IP_ADDRESS} '${prepareExecute}'"
          }
        }
      }
    }
    stage('Deploy Streamlit App') {
      steps {
        script {
          echo 'Deploying Streamlit app...'
          def deployLine = 'sudo systemctl start bca_streamlit'
          sshagent(['spectra-ec2']) {
            sh "scp -o StrictHostKeyChecking=no ubuntu@${EC2_IP_ADDRESS}:/home/ubuntu/"
            sh "ssh -o StrictHostKeyChecking=no ubuntu@${EC2_IP_ADDRESS} '${deployLine}'"
          }
        }
      }
    }
  }
}