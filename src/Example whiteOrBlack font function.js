       var whiteOrBlack  = function (hex) {
          if (hex.indexOf('#') === 0) {
              hex = hex.slice(1);
          }
          // convert 3-digit hex to 6-digits.
          if (hex.length === 3) {
              hex = hex[0] + hex[0] + hex[1] + hex[1] + hex[2] + hex[2];
          }
          if (hex.length !== 6) {
              throw new Error('Invalid HEX color.');
          }
          var r = parseInt(hex.slice(0, 2), 16),
              g = parseInt(hex.slice(2, 4), 16),
              b = parseInt(hex.slice(4, 6), 16);
              
          if((r+b+g)/255/3 > 0.5){
            return("#000000");
          }else{
            return("#FFFFFF");
          }
      };
