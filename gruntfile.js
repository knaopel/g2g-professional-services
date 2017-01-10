module.exports = function (grunt) {
    var paths = {
        src: '.\\src',
        dest: '.\\dist'
    };

    grunt.initConfig({
        pkg: grunt.file.readJSON('package.json'),
        meta: {
            banner: '/*\n' +
            '* <%= pkg.name %> - version <%= pkg.version %>\n' +
            '*\n' +
            '* <%= pkg.description %>\n' +
            '*\n' +
            '* Copyright <%= grunt.template.today("yyyy") %> <%= pkg.author %>\n' +
            '*\n' +
            '* Date: <%= grunt.template.today("yyyy-mm-dd, h:MM:ss TT") %>\n' +
            '*/',
            srcPath: '.\\src',
            nodeModPath: '.\\node_modules',
            distPath: '.\\dist'
        },

        usebanner: {
            dev: {
                options: {
                    position: 'top',
                    banner: '<%= meta.banner %>\n\n' +
                    '/***************************\n' +
                    ' **  FOR DEBUGGING ONLY   **\n' +
                    ' **************************/\n'
                },
                files: {
                    src: ['<%= meta.distPath %>\\*.js']
                }
            }
        },

        // CLEAN
        //-----------------------------------------------
        // https://npmjs.org/package/grunt-contrib-clean
        clean: {
            options: {
                force: true
            },
            files: ['<%= meta.distPath %>\\*.*']
        },

        // BUMP
        //-----------------------------------------------
        // https://npmjs.org/package/grunt-bump
        bump: {
            options: {
                files: ['package.json'],
                commit: false,
                createTag: false,
                push: false,
                prereleaseName: 'rc'
            }
        },

        // CONCAT
        //-----------------------------------------------
        // https://npmjs.org/package/grunt-contrib-concat
        concat: {
            dev: {
                files: {
                    '<%= meta.distPath %>\\G2G.Apps.ContentSection.Extensions.js': [
                        '<%= meta.nodeModPath %>\\file-saver\\FileSaver.js',
                        '<%= meta.nodeModPath %>\\xlsx\\dist\\xlsx.core.min.js',
                        '<%= meta.srcPath %>\\G2G.Apps.ContentSection.Extensions.js'

                    ]
                }
            }
        },

        // COPY
        //-----------------------------------------------
        // https://npmjs.org/package/grunt-contrib-copy
        copy: {
            txt: {
                files: [{
                    expand: true,
                    cwd: '<%= meta.srcPath %>\\',
                    src: ['*.html', '*.css'],
                    dest: '<%= meta.distPath %>\\'
                }]
            }
        }
    });

    grunt.loadNpmTasks('grunt-banner');
    grunt.loadNpmTasks('grunt-bump');
    grunt.loadNpmTasks('grunt-contrib-clean');
    grunt.loadNpmTasks('grunt-contrib-concat');
    grunt.loadNpmTasks('grunt-contrib-copy');

    grunt.registerTask('default', ['clean', 'copy', 'concat', 'usebanner:dev']);
};