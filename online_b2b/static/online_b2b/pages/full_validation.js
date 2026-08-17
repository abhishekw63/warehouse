/* online_b2b/full_validation.html — page script (separated from template). */
document.querySelectorAll('.fv-tab').forEach(function(t){
    t.addEventListener('click',function(){
      document.querySelectorAll('.fv-tab').forEach(function(x){x.classList.remove('on');});
      document.querySelectorAll('.fv-pane').forEach(function(x){x.classList.remove('on');});
      t.classList.add('on');
      var p=document.querySelector('.fv-pane[data-fvpane="'+t.getAttribute('data-fvtab')+'"]');
      if(p)p.classList.add('on');
    });
  });
