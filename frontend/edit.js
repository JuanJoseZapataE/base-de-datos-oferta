// edit.js - manejar edición de una ficha

function qparam(name){ const url = new URL(window.location.href); return url.searchParams.get(name); }

const fieldsOrder = [
  'cod_regional','regional','cod_municipio','municipio','cod_centro','centro_formacion',
  'cod_programa','denominacion_programa','cod_ficha','estado_ficha','jornada','nivel_formacion',
  'cupo','inscritos_primera_opcion','inscritos_segunda_opcion','oferta','tipo','perfil_ingreso','periodo'
];

async function setStatus(s){ document.getElementById('status').innerText = s; }

async function loadFicha(){
  const id = qparam('cod_ficha');
  if(!id){ setStatus('No se recibió cod_ficha en la URL'); return; }
  setStatus('Cargando ficha...');
  try{
    const base = document.getElementById('apiBase') ? document.getElementById('apiBase').value.replace(/\/+$/,'') : 'http://127.0.0.1:8000';
    const resp = await fetch(base + '/fichas/' + encodeURIComponent(id));
    if(!resp.ok){ const d = await resp.json(); throw new Error(d.detail || resp.statusText); }
    const data = await resp.json();
    populateForm(data);
    setStatus('');
  }catch(e){ setStatus('Error: '+e.message); console.error(e); }
}

function makeInput(key, value){
  const wrap = document.createElement('div');
  wrap.className = 'col-md-6';
  const label = document.createElement('label');
  label.className = 'form-label';
  label.innerText = key.replace(/_/g,' ');
  const input = document.createElement('input');
  input.className = 'form-control';
  input.id = 'f_' + key;
  input.name = key;
  input.value = value === null || value === undefined ? '' : value;
  wrap.appendChild(label);
  wrap.appendChild(input);
  return wrap;
}

function populateForm(data){
  const container = document.getElementById('fields');
  container.innerHTML = '';
  fieldsOrder.forEach(k=>{
    const v = data[k] === undefined ? '' : data[k];
    const fld = makeInput(k, v);
    if(k === 'cod_ficha') fld.querySelector('input').readOnly = true;
    container.appendChild(fld);
  });
}

async function saveFicha(e){
  e.preventDefault();
  const id = qparam('cod_ficha');
  if(!id) return alert('cod_ficha faltante');
  const body = {};
  fieldsOrder.forEach(k=>{
    const val = document.getElementById('f_' + k).value;
    if(val !== ''){
      // tipos simples: convertir a número cuando corresponda
      if(['cod_regional','cod_municipio','cod_centro','cod_programa','cod_ficha','cupo','inscritos_primera_opcion','inscritos_segunda_opcion','periodo'].includes(k)){
        body[k] = Number(val);
        if(Number.isNaN(body[k])) body[k] = null;
      } else {
        body[k] = val;
      }
    } else {
      body[k] = null;
    }
  });
  // No enviar cod_ficha en el payload (pk), backend lo toma de la URL
  delete body.cod_ficha;
  const base = document.getElementById('apiBase') ? document.getElementById('apiBase').value.replace(/\/+$/,'') : 'http://127.0.0.1:8000';
  setStatus('Guardando...');
  try{
    const resp = await fetch(base + '/fichas/' + encodeURIComponent(id), { method: 'PUT', headers: {'Content-Type':'application/json'}, body: JSON.stringify(body) });
    const d = await resp.json();
    if(!resp.ok){ throw new Error(d.detail || resp.statusText); }
    setStatus('Guardado. Filas afectadas: '+(d.updated_rows||0));
    setTimeout(()=>{ window.location.href = 'index.html'; }, 700);
  }catch(e){ setStatus('Error guardando: '+e.message); console.error(e); }
}

document.addEventListener('DOMContentLoaded', ()=>{
  const form = document.getElementById('editForm');
  form.addEventListener('submit', saveFicha);
  // inject a hidden apiBase input so edit page uses same base as index if present
  const baseInput = document.createElement('input'); baseInput.type='hidden'; baseInput.id='apiBase'; baseInput.value = 'http://127.0.0.1:8000'; document.body.appendChild(baseInput);
  loadFicha();
});
