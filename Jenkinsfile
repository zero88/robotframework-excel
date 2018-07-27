library identifier: 'notifications@master', retriever: modernSCM(
  [$class: 'GitSCMSource',
   remote: 'https://github.com/zero-88/jenkins-pipeline-shared.git'])

@Library('emailNotifications') _

def env_dockers = ["python-2.7": ["python:2.7.14-alpine3.7", "py2"], "python-3.6": ["python:3.6.4-alpine3.7", "py3"]]
def docker_build = "python:3.6.4-alpine3.7"
def envs = ["python-2.7", "python-3.6"]
def analysis_dir = "py3-out"

def get_build_stage(docker_image) {
    return {
        docker.image(docker_image).inside {
            stage("${docker_image}") {
                echo "Running in ${docker_image}"
            }

            stage("Build") {
                sh "pip install -r requirements.txt . --upgrade"
                sh "python -m robot.libdoc -f html ExcelRobot/ ./docs/ExcelRobot.html"
            }

        }
    }
}

def get_test_stage(docker_image, out) {
    return {
        docker.image(docker_image).inside {
            stage("${docker_image}") {
                echo "Running in ${docker_image}"
                sh "mkdir -p ${out}/unit"
                sh "mkdir -p ${out}/uat"
                sh "mkdir -p ${out}/coverage"
            }

            stage("Unit Test") {
                sh "pip install -r requirements.test.txt . --upgrade"
                sh "coverage run --source ExcelRobot -m nose tests.unit -v --with-xunit --xunit-file=${out}/unit/nosetests.xml -s --debug=ExcelRobot"
            }

            stage("Acceptance Test") {
                sh "coverage run -a --source ExcelRobot -m robot.run -d ${out}/uat ./tests/acceptance"
            }

            stage("Coverage") {
                sh "coverage report -m"
                sh "coverage html -d ${out}/coverage/"
                sh "coverage xml -o ${out}/coverage/coverage.xml"
            }

        }
    }
}


pipeline {
    parameters {
        booleanParam(defaultValue: true, description: 'Execute pipeline?', name: 'GO')
    }
    agent any

    stages {
        stage ("Preconditions") {
            steps {
                script {
                    result = sh (script: "git log -1 | grep '.*\\[ci skip\\].*'", returnStatus: true)
                    if (result == 0) {
                        echo ("This build should be skipped. Aborting.")
                        GO = "false"
                    }
                    VERSION = sh(script: 'python ExcelRobot/version.py', returnStdout: true).trim()
                }
            }
        }

        stage("Build") {
            when {
                beforeAgent true
                expression { BRANCH_NAME ==~ /^master|(feature|bugfix)\/.*/ }
                expression { return GO != "false" }
                tag "v*"
            }
            steps {
                script {
                    def build_stages = [:]
                    envs.each {
                        build_stages.put(it, get_build_stage(env_dockers.get(it)[0]))
                    }
                    parallel build_stages
                }
            }
        }

        stage("Test") {
            when {
                beforeAgent true
                expression { BRANCH_NAME ==~ /^master|(feature|bugfix)\/.*/ }
                expression { return GO != "false" }
                tag "v*"
            }
            steps {
                script {
                    def test_stages = [:]
                    envs.each {
                        test_stages.put(it, get_test_stage(env_dockers.get(it)[0], env_dockers.get(it)[1] + "-out"))
                    }
                    parallel test_stages
                }
            }
            post {
                always {
                    script {
                        envs.each {
                            def out = env_dockers.get(it)[1] + "-out"
                            junit "${out}/unit/nosetests.xml"
                            step([$class: "RobotPublisher",
                                        disableArchiveOutput: false,
                                        logFileName: "log.html",
                                        otherFiles: "",
                                        outputFileName: "output.xml",
                                        outputPath: "${out}/uat",
                                        passThreshold: 100,
                                        reportFileName: "report.html",
                                        unstableThreshold: 0])
                            zip archive: true, dir: "${out}", zipFile: "dist/test-${out}.zip"
                        }
                    }
                }
            }
        }

        stage("Analysis") {
            when {
                beforeAgent true
                expression { BRANCH_NAME ==~ /^master|(feature|bugfix)\/.*/ }
                expression { return GO != "false" }
                tag "v*"
            }
            agent {
                docker {
                    image "java:8-jre"
                    reuseNode true 
                }
            }
            steps {
                sh 'curl -L https://sonarsource.bintray.com/Distribution/sonar-scanner-cli/sonar-scanner-cli-3.2.0.1227.zip -o /tmp/sonar-scanner.zip'
                sh 'unzip /tmp/sonar-scanner.zip -d /tmp/'
                script {
                    withCredentials([string(credentialsId: 'SONAR_TOKEN', variable: 'SONAR_TOKEN')]) {
                        sh "set +x"
                        sh "/tmp/sonar-scanner-3.2.0.1227/bin/sonar-scanner -Dsonar.projectVersion=${VERSION} -Dsonar.python.xunit.reportPath=${analysis_dir}/unit/nosetests.xml -Dsonar.python.coverage.reportPath=${analysis_dir}/coverage/coverage.xml -Dsonar.host.url=https://sonarcloud.io -Dsonar.login=${SONAR_TOKEN}"
                    }
                }
            }
        }

        stage("Release") {
            when { tag "v*" }
            agent {
                docker {
                    image "${docker_build}"
                    reuseNode true 
                }
            }
            environment {
                PYPI_CREDS = credentials('pypi-credentials')
            }
            steps {
                echo "Continous Delivery"
                sh 'pip install twine --upgrade'
                sh 'python setup.py sdist bdist_wheel'
                sh "set +x"
                sh "twine upload dist/*.whl dist/*.tar.gz --repository-url ${PYPI_URL} -u ${PYPI_CREDS_USR} -p ${PYPI_CREDS_PSW}"
            }
        }

    }

    post {
        success {
            script {
                echo "${BRANCH_NAME}"
                if (BRANCH_NAME ==~ /^master|(feature|bugfix)\/.*/ && GO != "false") {
                    sshagent (credentials: ['cibot']) {
                        sh 'git add docs/ExcelRobot.html'
                        sh 'git diff-index --quiet HEAD || git commit -m "[ci skip]"'
                        sh "git push origin HEAD:${BRANCH_NAME}"
                    }
                }
            }
        }
        always {
            emailNotifications VERSION
        }
    }
}