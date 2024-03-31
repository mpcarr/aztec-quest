module.exports = function(grunt) {
  grunt.initConfig({
      concat: {
        dist: {
          src: ['src/vpx/**/*.vbs', 'src/game/**/*.vbs', '!src/unittests/**/*.vbs', '!src/**/*.test.vbs'],
          dest: 'dest/tablescript.vbs',
        },
        state: {
          src: ['src/game/**/*.vbs'],
          dest: 'dest/tablescript-state.vbs',
        },
        tests:{
          src: ['src/unittests/vbsUnit.vbs', 'src/unittests/mocks/**/*.vbs', 'dest/tablescript-state.vbs', 'src/unittests/tests-init.vbs','src/**/*.test.vbs', 'src/unittests/tests-report.vbs'],
          dest: 'dest/tests.vbs',
        }
      },
      exec: {
          tests: 'cscript dest/tests.vbs'
      },
      watch: {
        scripts: {
            files: 'src/**/*.vbs',
            tasks: ['concat'],
            //tasks: ['concat', 'exec:tests'],
            options: {  
                atBegin: true
            }
        }
      }
  
    });
  
    grunt.loadNpmTasks('grunt-contrib-concat');
    grunt.loadNpmTasks('grunt-contrib-watch');
    //grunt.loadNpmTasks('grunt-exec');

  
    // Default task(s).
    grunt.registerTask('default', ['concat']);
  
  };