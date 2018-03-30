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
                        sh 'pybot -d ./out ./tests/acceptance'
                        step([$class: 'RobotPublisher',
                                disableArchiveOutput: false,
                                logFileName: "log.html",
                                otherFiles: '',
                                outputFileName: "output.xml",
                                outputPath: '${env.WORKSPACE}/out',
                                passThreshold: 100,
                                reportFileName: "report.html",
                                unstableThreshold: 0])
                    }
                )
            }
        }

    }

    post {
        always {
            junit "out/nosetests.xml"
        }
        failure {
            emailext (
                to: "${env.DEFAULT_RECIPIENTS}",
                subject: "${env.DEFAULT_SUBJECT}",
                body: "${env.DEFAULT_CONTENT}",
                attachLog: true,
            )
        }
    }
}