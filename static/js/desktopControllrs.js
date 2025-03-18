/*
Copyright 2025 Adobe. All rights reserved.
This file is licensed to you under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License. You may obtain a copy
of the License at http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software distributed under
the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR REPRESENTATIONS
OF ANY KIND, either express or implied. See the License for the specific language
governing permissions and limitations under the License.
*/
$(document).ready(function() {
    const mydiv = $('#my-div');
      mydiv.hide();
      $("input[name$='installation']").click(function() {
          var test = $(this).val();
  
          $("div.desc").hide();
          $("#installation" + test).show();
      });
      
      $('#product').on('change', function(){
      var name= $('#product option:selected').val();
      if(name!=""){
          $(".content .value").html("Path : /Applications/"+name+"/"+name+".app");
          $("#productName").val("/Applications/"+name+"/"+name+".app");
          $("#productName1").val("/Applications/"+name+"/"+name+".app");
          const span = document.getElementById("my-span"); 
          // const mydiv= document.getElementById("my-div");
          const mydiv = $('#my-div');
  
          $.ajax({
              url: '/check_file?filepath='+"/Applications/"+name+"/"+name+".app",
              method: 'GET',
              data: $(this).serialize(),
              success: function(response) {
                  if (response.exists) {
                      // $('#my-span').text('File exists');
                      span.innerHTML='';
                      // mydiv.style.display=none;
                      if (mydiv) {
                        mydiv.hide();
        } else {
          console.log('');
        }
                  } else {
                    mydiv.show();
                      span.innerHTML='App is not present in the default location. Please Use Custom Installation !';
                  }
              }
          });
      }
      else if(name==""){
        const span = document.getElementById("my-span");
        const mydiv = $('#my-div');
        $(".content .value").html("");
        
          $("#productName").val("");
          $("#productName1").val("");
          span.innerHTML='';
          mydiv.hide();
      }
  });
  
  
  
  });

  (function($) {
    var CheckboxDropdown = function(el) {
      var _this = this;
      this.isOpen = false;
      this.areAllChecked = false;
      this.$el = $(el);
      this.$label = this.$el.find('.dropdown-label');
      this.$checkAll = this.$el.find('[data-toggle="check-all"]').first();
      this.$inputs = this.$el.find('[type="checkbox"]');
      
      this.onCheckBox();
      
      this.$label.on('click', function(e) {
        e.preventDefault();
        _this.toggleOpen();
      });
      
      this.$checkAll.on('click', function(e) {
        e.preventDefault();
        _this.onCheckAll();
      });
      
      this.$inputs.on('change', function(e) {
        _this.onCheckBox();
      });
    };
    
    CheckboxDropdown.prototype.onCheckBox = function() {
      this.updateStatus();
    };
    
    CheckboxDropdown.prototype.updateStatus = function() {
      var checked = this.$el.find(':checked');
      
      this.areAllChecked = false;
      this.$checkAll.html('Check All');
      
      if(checked.length <= 0) {
        this.$label.html('<i class="fa fa-language"></i> Select IME Locales');
      }
      else if(checked.length === 1) {
        this.$label.html(checked.parent('label').text());
      }
      else if(checked.length === this.$inputs.length) {
        this.$label.html('All Selected');
        this.areAllChecked = true;
        this.$checkAll.html('Uncheck All');
      }
      else {
        this.$label.html(checked.length + ' Selected');
      }
    };
    
    CheckboxDropdown.prototype.onCheckAll = function(checkAll) {
      if(!this.areAllChecked || checkAll) {
        this.areAllChecked = true;
        this.$checkAll.html('Uncheck All');
        this.$inputs.prop('checked', true);
      }
      else {
        this.areAllChecked = false;
        this.$checkAll.html('Check All');
        this.$inputs.prop('checked', false);
      }
      
      this.updateStatus();
    };
    
    CheckboxDropdown.prototype.toggleOpen = function(forceOpen) {
      var _this = this;
      
      if(!this.isOpen || forceOpen) {
         this.isOpen = true;
         this.$el.addClass('on');
        $(document).on('click', function(e) {
          if(!$(e.target).closest('[data-control]').length) {
           _this.toggleOpen();
          }
        });
      }
      else {
        this.isOpen = false;
        this.$el.removeClass('on');
        $(document).off('click');
      }
    };
    
    var checkboxesDropdowns = document.querySelectorAll('[data-control="checkbox-dropdown"]');
    for(var i = 0, length = checkboxesDropdowns.length; i < length; i++) {
      new CheckboxDropdown(checkboxesDropdowns[i]);
    }
    })(jQuery);