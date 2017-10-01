/**
 * Display SP-Wiki articles in categories
 */

'use strict';

var WIKICATEGORIES = WIKICATEGORIES || {
    WEB_SUBWEB_Name: 'websites/testweb', //empty if not used in subweb
    WIKI_CATEGORY_COLUMN_NAME: 'Category',
    CATEGORY_FOR_ITEMS_WITH_NO_CATEGORY: 'Unsorted',
    CATEGORY_SORT_COLUMN_NAME: 'Sort',
    CATEGORY_BACKGROUNDCOLOR_COLUMN_NAME: 'BackgroundColor',
    CATEGORY_WIDTH_COLUMN_NAME: 'Width',
    CATEGORY_TEXTCOLOR_COLUMN_NAME: 'TextColor',
    LIST_NAME_WIKI: 'Wiki',
    LIST_NAME_WIKI_CATEGORIES: "Wiki Categories",

    init: function () {
        var wikiCategories = this.getListItems(this.LIST_NAME_WIKI_CATEGORIES, '?$orderby=(' + this.CATEGORY_SORT_COLUMN_NAME + ')');
        var that = this;
        $.each(wikiCategories, function (index, value) {
            var title_orig = value.Title;
            var title = that.removeSpecialChars(title_orig);
            
            var id = 'wiki_category_' + title;
            var idBoxHeader = id + '_boxheader'
            var boxHeader = '<div id="' + idBoxHeader + '" class="wiki_category_box_header">' + title_orig + '</div>';
            var boxContent = '<div class="wiki_category_box_content"></div>';
            var bgColor = value[that.CATEGORY_BACKGROUNDCOLOR_COLUMN_NAME];
            var txtColor = value[that.CATEGORY_TEXTCOLOR_COLUMN_NAME];
            var boxWidth = value[that.CATEGORY_WIDTH_COLUMN_NAME];
            $("#wiki_categories").append('<div id="' + id + '" class="wiki_category_box">' + boxHeader + boxContent + '</div>');

            if (txtColor && txtColor.length > 0) {
                $("#" + idBoxHeader).css('color', txtColor);
            }

            if (bgColor && bgColor.length > 0) {
                $("#" + idBoxHeader).css('background-color', bgColor);
                $("#" + idBoxHeader).css('border-color', bgColor);
                $("#" + id).css('border-color', bgColor);
            }

            if (boxWidth && boxWidth > 0) {
                $("#" + id).css('width', boxWidth);
            }
        });

        this.retrieveListItems(this.LIST_NAME_WIKI);
    },

    getListItems: function(listName, params) {

        var result = '';
        var usedParams = ''

        if (typeof params != 'undefined') {
            usedParams = params;
        }

        var url = '';
        if (this.WEB_SUBWEB_Name) {
            url = "/" + this.WEB_SUBWEB_Name + "/_api/web/lists/GetByTitle('" + listName + "')/items" + usedParams;
        } else {
            url = "/_api/web/lists/GetByTitle('" + listName + "')/items" + usedParams;
        }

        $.ajax({
            url: url,
            async: false,
            type: "GET",
            headers: {
                "accept": "application/json;odata=verbose"
            },
        }).success(function (data) {
            result = data.d.results;
        }).fail(function (jqXHR, textStatus, errorThrown) {
            alert(errorThrown + ' - ' + textStatus);
        });

        return result;
    },

    appendToCategoryBoxContent: function(className, oListItem) {
        var currWikiItenName = oListItem.get_item('FileLeafRef').toString().replace('.aspx', '')
        var currWikiItemUrl = oListItem.get_item('FileRef');
        var currWikiItemLink = '<a href="' + currWikiItemUrl + '">' + currWikiItenName + '</a>';
        $('#' + className + ' > .wiki_category_box_content').append(currWikiItemLink + '<br >');
        
    },

    retrieveListItems: function(listName) {
        var clientContext = new SP.ClientContext();
        var oList = clientContext.get_web().get_lists().getByTitle(listName);
        var camlQuery = new SP.CamlQuery();

        this.collListItem = oList.getItems(camlQuery);
        clientContext.load(this.collListItem);
        clientContext.executeQueryAsync(
            Function.createDelegate(this, this.onQuerySucceeded),
            Function.createDelegate(this, this.onQueryFailed)
        );
    },

    onQuerySucceeded: function(sender, args) {
    
        var listItemEnumerator = this.collListItem.getEnumerator();
        var that = this;
        while (listItemEnumerator.moveNext()) {
            var oListItem = listItemEnumerator.get_current();
            var cats = oListItem.get_item(that.WIKI_CATEGORY_COLUMN_NAME);
            if (cats.length > 0) {
                $.each(cats, function (index, value) {
                    var className = 'wiki_category_' + that.removeSpecialChars(value.get_lookupValue());
                    that.appendToCategoryBoxContent(className, oListItem);
                });
            } else {
                //Items without category
                var className = 'wiki_category_' + this.CATEGORY_FOR_ITEMS_WITH_NO_CATEGORY;
                that.appendToCategoryBoxContent(className, oListItem);
            }
        }
    },

    onQueryFailed: function(sender, args) {
        alert('Request failed. ' + args.get_message() +
            '\n' + args.get_stackTrace());
    },

    removeSpecialChars: function (value) {
        if (typeof(value) != "undefined") {
            return value.replace(/[^A-Za-z0-9\-_]/g, '_');
        }
    }
};

$(function () {
    WIKICATEGORIES.init();
});
