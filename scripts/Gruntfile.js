module.exports = function (grunt) {
  grunt.initConfig({
    concat: {
      vpx: {
        src: ['src/vpx/**/*.vbs', 'src/game/**/*.vbs', '!src/unittests/**/*.vbs', '!src/**/*.test.vbs', '!src/**/*-mpf.vbs'],
        dest: 'dest/vpx/tablescript.vbs',
      },
      mpf: {
        src: ['src/vpx/**/*.vbs', 'src/game/**/*.vbs', '!src/unittests/**/*.vbs', '!src/**/*.test.vbs', '!src/**/*-vpx.vbs'],
        dest: 'dest/mpf/tablescript.vbs',
      },
    },
    exec: {
      tests: 'cscript dest/tests.vbs'
    },
    watch: {
      vpx: {
        files: 'src/**/*.vbs',
        tasks: ['concat:vpx'],
        //tasks: ['concat', 'exec:tests'],
        options: {
          atBegin: true
        }
      },
      mpf: {
        files: 'src/**/*.vbs',
        tasks: ['concat:mpf'],
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