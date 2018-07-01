def env_dockers = ["python-2.7": ["python:2.7.14-alpine3.7", "py2"], "python-3.6": ["python:3.6.4-alpine3.7", "py3"]]
def docker_build = "python:3.6.4-alpine3.7"
def envs = ["python-2.7", "python-3.6"]

def get_build_stage(docker_image, out) {
    return {
        docker.image(docker_image).inside {
            stage("${docker_image}") {
                echo "Running in ${docker_image}"
            }

            stage("Build") {
                sh "pip install . --upgrade"
                sh "python -m robot.libdoc -f html ExcelRobot/ ./docs/ExcelRobot.html"
            }

            stage("Unit Test") {
                sh "mkdir -p ${out}"
                sh "nosetests tests.unit -v --with-xunit --xunit-file=${out}/nosetests.xml -s --debug=ExcelRobot"
            }

            stage("Acceptance Test") {
                sh "pybot -d ${out} ./tests/acceptance"
            }
        }
    }
}


pipeline {
    agent any

    stages {

        stage("CI") {
            steps {
                echo "Continous Integration"
                script {
                    def build_stages = [:]
                    envs.each {
                        def docker_image = env_dockers.get(it)[0]
                        def out = env_dockers.get(it)[1] + "-out"
                        build_stages.put(it, get_build_stage(docker_image, out))
                    }
                    parallel build_stages
                }
            }
        }

        stage("CD") {
            agent { docker "${docker_build}" }
            when { tag "v*" }
            steps {
                echo "Continous Delivery"
                sh 'pip install twine --upgrade'
                sh 'python setup.py sdist bdist_wheel'
                sh 'twine upload dist/*.whl dist/*.tar.gz'
            }
        }

    }

    post {
        always {
            script {
                envs.each {
                    def out = env_dockers.get(it)[1] + "-out"
                    echo "${out}"
                    zip archive: true, dir: "${out}", glob: "*.xml,*.html", zipFile: "dist/test-${out}.zip"
                    junit "${out}/nosetests.xml"
                    step([$class: "RobotPublisher",
                                disableArchiveOutput: false,
                                logFileName: "log.html",
                                otherFiles: "",
                                outputFileName: "output.xml",
                                outputPath: "${out}",
                                passThreshold: 100,
                                reportFileName: "report.html",
                                unstableThreshold: 0])
                    archiveArtifacts artifacts: "${out}/**/*.xml,${out}/**/*.html", fingerprint: true
                }
            }
        }
        failure {
            script {
                def committerEmail = sh (
                    script: 'git --no-pager show -s --format=\'%ae\'',
                    returnStdout: true
                ).trim()
                def committer = sh (
                    script: 'git --no-pager show -s --format=\'%an\'',
                    returnStdout: true
                ).trim()
                def content = """
                    - Job Name: ${env.JOB_NAME}
                    - Build URL: ${env.BUILD_URL}
                    - Changes:
                        - ${committer} <${committerEmail}>
                        - ${env.GIT_COMMIT}
                        - ${env.GIT_BRANCH}
                        - ${env.GIT_URL}
                """
                emailext (
                    recipientProviders: [[$class: "DevelopersRecipientProvider"]],
                    subject: "[Jenkins] ${env.JOB_NAME}-#${env.BUILD_NUMBER} [${currentBuild.result}]",
                    body: "${content}",
                    attachLog: true,
                    compressLog: true
                )
            }
            
        }
    }
}