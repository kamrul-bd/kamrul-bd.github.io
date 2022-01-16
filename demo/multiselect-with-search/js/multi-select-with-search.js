/*
* MultiSelectWithSearch v0.1.0
* Copyright (c) 2021 Devnet limited
*
* This program is totally free. It is combination of multi-select and quicksearch.
* It comes without any warranty, to
* the extent permitted by applicable law. You can redistribute it
* and/or modify it
*/
!function ($) {

  $.fn.multiSelectWithSearch = function (options) {

    var allOptions = $.extend({
            selectableHeader: `<input type='text' class='search-input' autocomplete='off' placeholder='${(options && options.selectableSearchPlaceHolder ? options.selectableSearchPlaceHolder : "Type please")}'>`,
            selectionHeader: `<input type='text' class='search-input' autocomplete='off' placeholder='${(options && options.selectionSearchPlaceHolder ? options.selectionSearchPlaceHolder : "Type please")}'>`,
            afterInit: function(ms){
                var that = this,
                    $selectableSearch = that.$selectableUl.prev(),
                    $selectionSearch = that.$selectionUl.prev(),
                    selectableSearchString = '#'+that.$container.attr('id')+' .ms-elem-selectable:not(.ms-selected)',
                    selectionSearchString = '#'+that.$container.attr('id')+' .ms-elem-selection.ms-selected';

                that.qs1 = $selectableSearch.quicksearch(selectableSearchString)
                    .on('keydown', function(e){
                        if (e.which === 40){
                            that.$selectableUl.focus();
                            return false;
                        }
                    });

                that.qs2 = $selectionSearch.quicksearch(selectionSearchString)
                    .on('keydown', function(e){
                        if (e.which == 40){
                            that.$selectionUl.focus();
                            return false;
                        }
                    });
            },
            afterSelect: function(){
                this.qs1.cache();
                this.qs2.cache();
            },
            afterDeselect: function(){
                this.qs1.cache();
                this.qs2.cache();
            },
        }, options);


        $(this['selector']).multiSelect(allOptions);
    };

} (window.jQuery);