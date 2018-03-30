pipeline {
    agent { 
        docker { image 'python:3.6.4-alpine3.7' }
    }

    stages {
        stage('Pre-build') {
            steps {
                checkout scm
            }
        }

        stage('Build') {

            steps {
                sh 'pip install . --upgrade'
                sh 'python -m robot.libdoc -f html ExcelRobot/ ./docs/ExcelRobot.html'
            }
            
        }

        stage('Test') {

            steps {
                parallel (
                    'unit': { 
                        sh 'mkdir -p ./out/'
                        sh 'nosetests tests.unit -v --with-xunit --xunit-file=./out/nosetests.xml -s --debug=ExcelRobot'
                    },
                    'acceptance': { 
                        sh 'python -m robot.libdoc -f html ExcelRobot/ ./docs/ExcelRobot.html'
                        step([$class: 'RobotPublisher',
                                disableArchiveOutput: false,
                                logFileName: 'log.html',
                                otherFiles: '',
                                outputFileName: 'output.xml',
                                outputPath: '.',
                                passThreshold: 100,
                                reportFileName: 'report.html',
                                unstableThreshold: 0])
                    }
                )
            }
        }

    }

    post {
        always {
            junit './out/nosetests.xml'
        }
        failure {
            emailext (
                to: 'sontt246@gmail.com',
                subject: "${env.JOB_NAME} #${env.BUILD_NUMBER} [${currentBuild.result}]",
                body: "Build URL: ${env.BUILD_URL}.\n\n",
                attachLog: true,
            )
        }
    }
}