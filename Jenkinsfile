pipeline {
    agent any

    stages {
        stage('Pre-build') {
            steps {
                checkout scm
            }
        }

        stage('Build') {
            // parallel {
            //     stage('py2') {
            //         agent { 
            //             docker { image 'python:2.7.14-alpine3.7' } 
            //         }

            //         stage('Build') {
            //             steps {
            //                 sh 'pip install . --upgrade'
            //                 sh 'python -m robot.libdoc -f html ExcelRobot/ ./docs/ExcelRobot.html'
            //             }
            //         }

            //         stage('Test') {
            //             steps {
            //                 parallel (
            //                     'unit tests': { sh 'nosetests tests.unit -v --with-xunit --xunit-file=./out/nosetests.xml -s --debug=ExcelRobot' },
            //                     'acceptance tests': { sh 'python -m robot.libdoc -f html ExcelRobot/ ./docs/ExcelRobot.html' }
            //                 )
            //             }
            //         }
            //     }

                // stage('py3') {
                    agent { 
                        docker { image 'python:3.6.4-alpine3.7' } 
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
                                'unit tests': { sh 'nosetests tests.unit -v --with-xunit --xunit-file=./out/nosetests.xml -s --debug=ExcelRobot' },
                                'acceptance tests': { sh 'python -m robot.libdoc -f html ExcelRobot/ ./docs/ExcelRobot.html' }
                            )
                        }
                    }
                // }
            // }
        }

    }

    post {
        failure {
            mail to: 'sontt246@gmail.com', subject: 'The Pipeline failed :(', body: "${env.BUILD_URL}"
        }
    }
}