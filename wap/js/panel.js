(function(c,b,d){var e=b.fx.cssPrefix,a=b.fx.transitionEnd;c.define("Panel",{options:{contentWrap:"",scrollMode:"follow",display:"push",position:"right",dismissible:true,swipeClose:true},_init:function(){var g=this,f=g._options;g.on("ready",function(){g.displayFn=g._setDisplay();g.$contentWrap.addClass("ui-panel-animate");g.$el.on(a,b.proxy(g._eventHandler,g)).hide();f.dismissible&&g.$panelMask.hide().on("click",b.proxy(g._eventHandler,g));f.scrollMode!=="follow"&&b(window).on("scrollStop",b.proxy(g._eventHandler,g));b(window).on("ortchange",b.proxy(g._eventHandler,g))})},_create:function(){if(this._options.setup){var h=this,g=h._options,f=h.$el.addClass("ui-panel ui-panel-"+g.position);h.panelWidth=f.width()||0;h.$contentWrap=b(g.contentWrap||f.next());g.dismissible&&(h.$panelMask=b('<div class="ui-panel-dismiss"></div>').width(document.body.clientWidth-f.width()).appendTo("body")||null)}else{throw new Error("panel\u7ec4\u4ef6\u4e0d\u652f\u6301create\u6a21\u5f0f\uff0c\u8bf7\u4f7f\u7528setup\u6a21\u5f0f")}},_setDisplay:function(){var j=this,l=j.$el,g=j.$contentWrap,f=e+"transform",m=j._transDisplayToPos(),k={},i,h;b.each(["push","overlay","reveal"],function(n,o){k[o]=function(q,r,p){i=m[o].panel,h=m[o].cont;l.css(f,"translate3d("+j._transDirectionToPos(r,i[q])+"px,0,0)");if(!p){g.css(f,"translate3d("+j._transDirectionToPos(r,h[q])+"px,0,0)");j.maskTimer=setTimeout(function(){j.$panelMask&&j.$panelMask.css(r,l.width()).toggle(q)},400)}return j}});return k},_initPanelPos:function(f,g){this.displayFn[f](0,g,true);this.$el.get(0).clientLeft;return this},_transDirectionToPos:function(g,f){return g==="left"?f:-f},_transDisplayToPos:function(){var f=this,g=f.panelWidth;return{push:{panel:[-g,0],cont:[0,g]},overlay:{panel:[-g,0],cont:[0,0]},reveal:{panel:[0,0],cont:[0,g]}}},_setShow:function(h,i,o){var n=this,g=n._options,m=h?"open":"close",p=b.Event("before"+m),l=h!==n.state(),f=h?"on":"off",j=h?b.proxy(n._eventHandler,n):n._eventHandler,k=i||g.display,q=o||g.position;n.trigger(p,[i,o]);if(p.isDefaultPrevented()){return n}if(l){n._dealState(h,k,q);n.displayFn[k](n.isOpen=Number(h),q);g.swipeClose&&n.$el[f](b.camelCase("swipe-"+q),j);g.display=k,g.position=q}return n},_dealState:function(g,i,l){var j=this,f=j._options,m=j.$el,h=j.$contentWrap,n="ui-panel-"+i+" ui-panel-"+l,k="ui-panel-"+f.display+" ui-panel-"+f.position+" ui-panel-animate";if(g){m.removeClass(k).addClass(n).show();f.scrollMode==="fix"&&m.css("top",b(window).scrollTop());j._initPanelPos(i,l);if(i==="reveal"){h.addClass("ui-panel-contentWrap").on(a,b.proxy(j._eventHandler,j))}else{h.removeClass("ui-panel-contentWrap").off(a,b.proxy(j._eventHandler,j));m.addClass("ui-panel-animate")}j.$panelMask&&j.$panelMask.css({left:"auto",right:"auto",height:document.body.clientHeight})}return j},_eventHandler:function(i){var h=this,g=h._options,j=g.scrollMode,f=h.state()?"open":"close";switch(i.type){case"click":case"swipeLeft":case"swipeRight":h.close();break;case"scrollStop":j==="fix"?h.$el.css("top",b(window).scrollTop()):h.close();break;case a:h.trigger(f,[g.display,g.position]);break;case"ortchange":h.$panelMask&&h.$panelMask.css("height",document.body.clientHeight);j==="fix"&&h.$el.css("top",b(window).scrollTop());break}},open:function(g,f){return this._setShow(true,g,f)},close:function(){return this._setShow(false)},toggle:function(g,f){return this[this.isOpen?"close":"open"](g,f)},state:function(){return !!this.isOpen},destroy:function(){this.$panelMask&&this.$panelMask.off().remove();this.maskTimer&&clearTimeout(this.maskTimer);this.$contentWrap.removeClass("ui-panel-animate");b(window).off("scrollStop",this._eventHandler);b(window).off("ortchange",this._eventHandler);return this.$super("destroy")}})})(gmu,gmu.$);