import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import folium
from folium.plugins import MarkerCluster, Draw, MeasureControl
from folium.features import DivIcon
from branca.element import Template, MacroElement
import webbrowser
import json
import re

# ─── CONFIG ────────────────────────────────────────────────────────────────
CIRCLE_RADIUS_PIXELS = 6

STYLE_COLORS = {
    'No MR':            'green',
    'Comm MR':          'yellow',
    'Simple Power MR':  'orange',
    'Complex Power MR': 'red',
    'Deferred':         'gray',
    'Cannot Attach':    'black',
}
MISSING_COLOR    = 'purple'
WARNING_COLOR    = 'aqua'

PALETTE          = ['blue','darkblue','teal','navy','magenta','lime']
MULTI_COMP_COLOR = 'red'
MARKUP_COLORS = ['red', 'orange', 'yellow', 'green', 'blue', 'purple', 'aqua', 'lime', 'black']

# Path for optional markup file
markup_file = None

def convex_hull(points):
    pts = sorted(set(points))
    if len(pts) < 3:
        return pts
    def cross(o,a,b):
        return (a[0]-o[0])*(b[1]-o[1]) - (a[1]-o[1])*(b[0]-o[0])
    lower, upper = [], []
    for p in pts:
        while len(lower)>=2 and cross(lower[-2],lower[-1],p)<=0:
            lower.pop()
        lower.append(p)
    for p in reversed(pts):
        while len(upper)>=2 and cross(upper[-2],upper[-1],p)<=0:
            upper.pop()
        upper.append(p)
    return lower[:-1] + upper[:-1]

def build_map(df, level_cols, cost_cols, warning_cols,
              company_cols, view_mode, out_html="map.html",
              markup_path=None):
    df = df[~df['job_name'].str.contains(r'copy', case=False, na=False)]
    df = df[df['node_type'].astype(str).str.lower()=='pole']
    df = df.dropna(subset=['latitude','longitude'])
    df = df[df['scid'].astype(str).str.match(r'^\d+$', na=False)]

    if view_mode == 'Utility':
        df['base_job'] = df['job_name'].str.replace(r'\s*\bPLA\b', '', case=False, regex=True)
        has_pla = df[df['job_name'].str.contains(r'\bPLA\b', case=False, regex=True, na=False)]['base_job'].unique()
        mask_pla = df['base_job'].isin(has_pla)
        mask_is_pla = df['job_name'].str.contains(r'\bPLA\b', case=False, regex=True, na=False)
        df = df[(mask_pla & mask_is_pla) | (~mask_pla)]
        df = df.drop(columns=['base_job'])
    else:
        df = df[~df['job_name'].str.contains(r'\bPLA\b', case=False, regex=True, na=False)]

    if view_mode=='Utility':
        mask = df[company_cols].fillna('').astype(str).apply(
            lambda row: any(re.search(r'\bAT&T\b|\bCharter\b', cell, re.IGNORECASE)
                            for cell in row),
            axis=1
        )
        df = df[~mask]

    lats = df['latitude'].astype(float)
    lons = df['longitude'].astype(float)
    m = folium.Map(location=(lats.mean(), lons.mean()), zoom_start=10)
    feature_group = folium.FeatureGroup(name='Markups').add_to(m)
    MarkerCluster(disableClusteringAtZoom=7).add_to(m)
    Draw(export=True, feature_group=feature_group).add_to(m)
    MeasureControl().add_to(m)

    if view_mode=='Utility':
        def canon(v):
            v = v.strip()
            if v.lower() in ('magic valley elec coop','magic valley electric coop'):
                return 'Magic Valley Electric Coop'
            return v
        comps = []
        for _, r in df.iterrows():
            for c in company_cols:
                v = str(r.get(c,'')).strip()
                if v and v.lower()!='nan':
                    comps.append(canon(v))
        uniq = sorted(set(comps))
        comp_color = {
            'Magic Valley Electric Coop': '#98ff98',
            'AEP': 'orange',
            'Brownsville Public Utilities': 'aqua',
            'Central Bradford PA': 'yellow'
        }
        idx = 0
        for comp in uniq:
            if comp not in comp_color:
                comp_color[comp] = PALETTE[idx % len(PALETTE)]
                idx += 1

    def first_val(row, cols):
        for c in cols:
            v = row.get(c)
            if pd.notna(v) and str(v).strip():
                return str(v)
        return ''

    for _, row in df.iterrows():
        lat, lon = float(row['latitude']), float(row['longitude'])
        scid = row['scid']
        warnings = [str(row[c]) for c in warning_cols if pd.notna(row.get(c)) and str(row.get(c)).strip()]
        raw_comps = [str(row[c]) for c in company_cols if pd.notna(row.get(c)) and str(row.get(c)).strip()]

        if view_mode=='MR':
            companies = [raw_comps[0]] if raw_comps else []
        else:
            companies = [canon(v) for v in raw_comps]

        if view_mode=='MR':
            lvl  = first_val(row, level_cols)
            cost = first_val(row, cost_cols)
            tooltip = (
                f"<b>SCID:</b> {scid}<br>"
                f"<b>MR Level:</b> {lvl or '<i>None</i>'}<br>"
                f"<b>Cost:</b> {cost or '<i>None</i>'}"
            )
            if warnings:
                tooltip += "<br><b>Warnings:</b> " + "; ".join(warnings)
                folium.RegularPolygonMarker(
                    location=(lat,lon), number_of_sides=4, radius=CIRCLE_RADIUS_PIXELS,
                    color=WARNING_COLOR, fill_color=WARNING_COLOR, fill_opacity=0.9,
                    tooltip=tooltip, popup=row['job_name']
                ).add_to(m)
            else:
                color = STYLE_COLORS.get(lvl, MISSING_COLOR)
                folium.CircleMarker(
                    location=(lat,lon), radius=CIRCLE_RADIUS_PIXELS,
                    fill=True, fill_color=color, fill_opacity=0.8, stroke=False,
                    tooltip=tooltip, popup=row['job_name']
                ).add_to(m)
        else:
            tooltip = (
                f"<b>SCID:</b> {scid}<br>"
                f"<b>Companies:</b> {', '.join(companies) or '<i>None</i>'}"
            )
            if len(companies) > 1:
                folium.RegularPolygonMarker(
                    location=(lat,lon), number_of_sides=4, radius=CIRCLE_RADIUS_PIXELS,
                    color=MULTI_COMP_COLOR, fill_color=MULTI_COMP_COLOR, fill_opacity=0.9,
                    tooltip=tooltip
                ).add_to(m)
            elif len(companies) == 1:
                color = comp_color.get(companies[0], PALETTE[0])
                folium.CircleMarker(
                    location=(lat,lon), radius=CIRCLE_RADIUS_PIXELS,
                    fill=True, fill_color=color, fill_opacity=0.8, stroke=False,
                    tooltip=tooltip
                ).add_to(m)
            else:
                folium.map.Marker(
                    location=(lat,lon),
                    icon=DivIcon(
                        icon_size=(24,24), icon_anchor=(12,12),
                        html='<div style="color:red;font-size:16px;font-weight:bold;">X</div>'
                    ),
                    tooltip=tooltip
                ).add_to(m)

    hull_features = []
    for job, grp in df.groupby('job_name'):
        pts = [(float(r['longitude']), float(r['latitude'])) for _, r in grp.iterrows()]
        if len(pts) < 3:
            continue
        hull = convex_hull(pts)
        hull.append(hull[0])
        hull_features.append({
            'type': 'Feature',
            'properties': {
                'label': job,
                'color': 'black',
                'labelSize': 10,
                'isJobLabel': True,
                'opacity': 0.95
            },
            'geometry': {
                'type': 'Polygon',
                'coordinates': [[ [lon, lat] for lon, lat in hull ]]
            }
        })

    if hull_features:
        hull_gj = folium.GeoJson(
            {'type': 'FeatureCollection', 'features': hull_features},
            name='Job Hulls',
            style_function=lambda feat: {
                'color': feat['properties'].get('color', 'black'),
                'fill': False,
                'weight': 2,
                'opacity': feat['properties'].get('opacity', 0.95)
            }
        ).add_to(m)
        hull_macro = MacroElement()
        hull_macro._template = Template(
            """{% macro script(this, kwargs) %}
            {{ this.geojson }}.eachLayer(function(l){
                {{ this.fg }}.addLayer(l);
            });
            {% endmacro %}"""
        )
        hull_macro.geojson = hull_gj.get_name()
        hull_macro.fg = feature_group.get_name()
        m.get_root().add_child(hull_macro)

    m.fit_bounds([[lats.min(),lons.min()],[lats.max(),lons.max()]])

    if view_mode=='MR':
        entries = list(STYLE_COLORS.items()) + [
            ('Missing MR', MISSING_COLOR),
            ('Warnings', WARNING_COLOR)
        ]
        legend_html = "{% macro html(this,kwargs) %}<div style='position:fixed;bottom:50px;left:50px;width:180px;padding:10px;background:white;border:2px solid grey;z-index:9999;font-size:14px;'><strong>Legend</strong><br>"
        for name,col in entries:
            legend_html += f"<div><span style='background:{col};width:12px;height:12px;display:inline-block;border-radius:50%;margin-right:5px;'></span>{name}</div>"
        legend_html += "</div>{% endmacro %}"
    else:
        legend_html = "{% macro html(this,kwargs) %}<div style='position:fixed;bottom:50px;left:50px;width:240px;padding:10px;background:white;border:2px solid grey;z-index:9999;font-size:14px;'><strong>Legend</strong><br>"
        for comp,col in comp_color.items():
            legend_html += f"<div><span style='background:{col};width:12px;height:12px;display:inline-block;border-radius:50%;margin-right:5px;'></span>{comp}</div>"
        legend_html += "<div><span style='background:red;width:12px;height:12px;display:inline-block;margin-right:5px;'></span>Multiple Companies</div>"
        legend_html += "<div><span style='color:red;font-size:16px;font-weight:bold;display:inline-block;margin-right:5px;'>X</span>No Company</div>"
        legend_html += "</div>{% endmacro %}"

    legend = MacroElement()
    legend._template = Template(legend_html)
    m.get_root().add_child(legend)

    if markup_path:
        try:
            with open(markup_path, 'r') as f:
                data = json.load(f)
            gj = folium.GeoJson(
                data,
                name='Imported Markups',
                style_function=lambda feat: {
                    'color': feat['properties'].get('color', 'blue'),
                    'fillColor': feat['properties'].get('color', 'blue'),
                    'opacity': feat['properties'].get('opacity', 0.95),
                    'fillOpacity': feat['properties'].get('opacity', 0.95)
                }
            ).add_to(m)
            import_macro = MacroElement()
            import_macro._template = Template(
                """{% macro script(this, kwargs) %}
                {{ this.geojson }}.eachLayer(function(l){
                    {{ this.fg }}.addLayer(l);
                });
                {% endmacro %}"""
            )
            import_macro.geojson = gj.get_name()
            import_macro.fg = feature_group.get_name()
            m.get_root().add_child(import_macro)
        except Exception as e:
            messagebox.showwarning("Markup Error", f"Could not load markup file:\n{e}")

    # Custom JS for editing colors and labels
    script = Template("""
    {% macro script(this, kwargs) %}
    var COLORS = {{ this.colors | safe }};
    var container = {{ this.m.get_name() }}.getContainer();
    var style = document.createElement('style');
    style.innerHTML = '.hide-job-labels .job-label{display:none;}';
    document.head.appendChild(style);
    var styleExport = document.createElement('style');
    styleExport.innerHTML = '#export{position:absolute;top:auto!important;bottom:10px;right:10px;left:auto!important;width:auto;height:auto;}';
    document.head.appendChild(styleExport);
    var labelsHidden = false;
    var markupsHidden = false;
    function toggleMarkups(){
        markupsHidden = !markupsHidden;
        {{ this.feature_group.get_name() }}.eachLayer(function(layer){
            var props = (layer.feature && layer.feature.properties) || {};
            if(props.isJobLabel){
                return; // never toggle job hulls
            }
            if(markupsHidden){
                {{ this.m.get_name() }}.removeLayer(layer);
            } else {
                {{ this.m.get_name() }}.addLayer(layer);
            }
        });
    }

    function toggleLabels(){
        if(labelsHidden){
            L.DomUtil.removeClass(container, 'hide-job-labels');
        } else {
            L.DomUtil.addClass(container, 'hide-job-labels');
        }
        labelsHidden = !labelsHidden;
    }

    var labelToggle = L.control({position:'topright'});
    labelToggle.onAdd = function(map){
        var div = L.DomUtil.create('div','leaflet-bar leaflet-control');
        div.style.marginBottom = '10px';
        var a = L.DomUtil.create('a','',div);
        a.href = '#';
        a.innerHTML = 'J';
        a.title = 'Toggle Job Labels';
        L.DomEvent.on(a,'click',function(e){
            L.DomEvent.preventDefault(e);
            toggleLabels();
        });
        return div;
    };
    labelToggle.addTo({{ this.m.get_name() }});

    var markupToggle = L.control({position:'topright'});
    markupToggle.onAdd = function(map){
        var div = L.DomUtil.create('div','leaflet-bar leaflet-control');
        div.style.marginBottom = '10px';
        var a = L.DomUtil.create('a','',div);
        a.href = '#';
        a.innerHTML = 'M';
        a.title = 'Toggle Markups';
        L.DomEvent.on(a,'click',function(e){
            L.DomEvent.preventDefault(e);
            toggleMarkups();
        });
        return div;
    };
    markupToggle.addTo({{ this.m.get_name() }});
    function promptEdit(layer){
        var props = (layer.feature && layer.feature.properties) || {};
        var col = props.color || layer.options.color || 'blue';
        var label = props.label || '';
        var size = props.labelSize || 10;
        var opacity = props.opacity != null ? props.opacity : (layer.options.opacity != null ? layer.options.opacity : 0.95);
        var input = prompt('Color ('+COLORS.join(', ')+'):', col);
        if(input){ col = input; }
        input = prompt('Label (optional):', label);
        if(input !== null){ label = input; }
        input = prompt('Label size (px):', size);
        size = parseInt(input) || size;
        input = prompt('Opacity % (0-100):', Math.round(opacity*100));
        var o = parseFloat(input);
        if(!isNaN(o)) opacity = Math.max(0, Math.min(100, o))/100;
        if(layer.setStyle){
            layer.setStyle({color: col, fillColor: col, opacity: opacity, fillOpacity: opacity});
        }
        if(label){
            var cls = props.isJobLabel ? 'job-label' : '';
            layer.bindTooltip('<div style="font-size:'+size+'px">'+label+'</div>',{permanent:true, direction:'center', className:cls}).openTooltip();
        } else {
            layer.unbindTooltip();
        }
        layer.feature = layer.feature || {type:'Feature', properties:{}};
        layer.feature.properties.color = col;
        layer.feature.properties.label = label;
        layer.feature.properties.labelSize = size;
        layer.feature.properties.opacity = opacity;
        if(!{{ this.feature_group.get_name() }}.hasLayer(layer)){
            {{ this.feature_group.get_name() }}.addLayer(layer);
        }
    }

    function attach(layer){
        layer.on('contextmenu', function(){ promptEdit(layer); });
        var p = layer.feature && layer.feature.properties;
        if(p){
            if(layer.setStyle){
                var col = p.color || layer.options.color || 'blue';
                var op = p.opacity != null ? p.opacity : (layer.options.opacity != null ? layer.options.opacity : 0.95);
                layer.setStyle({color: col, fillColor: col, opacity: op, fillOpacity: op});
            }
            if(p.label){
                var s = p.labelSize || 10;
                var cls = p.isJobLabel ? 'job-label' : '';
                layer.bindTooltip('<div style="font-size:'+s+'px">'+p.label+'</div>',{permanent:true, direction:'center', className:cls}).openTooltip();
            }
            
        }
    }

    function traverse(layer){
        if(layer.eachLayer){
            layer.eachLayer(function(sub){ traverse(sub); });
        } else {
            attach(layer);
        }
    }

    {{ this.feature_group.get_name() }}.eachLayer(function(l){ traverse(l); });
    {{ this.feature_group.get_name() }}.on('layeradd', function(e){ traverse(e.layer); });

    {{ this.m.get_name() }}.on('draw:created', function(e){
        var layer = e.layer;
        promptEdit(layer);
        {{ this.feature_group.get_name() }}.addLayer(layer);
    });
    {% endmacro %}
    """)

    macro = MacroElement()
    macro._template = script
    macro.feature_group = feature_group
    macro.m = m
    macro.colors = json.dumps(MARKUP_COLORS)
    m.get_root().add_child(macro)

    folium.LayerControl().add_to(m)

    m.save(out_html)
    return out_html

def on_load_markup():
    global markup_file
    markup_file = filedialog.askopenfilename(
        title="Select Markup GeoJSON",
        filetypes=[("GeoJSON", "*.geojson *.json"), ("All Files", "*.*")]
    )


def on_generate():
    path = filedialog.askopenfilename(title="Select XLSX", filetypes=[("Excel","*.xlsx *.xls")])
    if not path:
        return
    try:
        df = pd.read_excel(path, engine='openpyxl')
    except Exception as e:
        return messagebox.showerror("Read Error", str(e))

    base = {'job_name','latitude','longitude','node_type','scid'}
    if not base.issubset(df.columns):
        return messagebox.showerror("Missing Columns", "Need: " + ", ".join(base))

    level_cols   = [c for c in df.columns if c.lower().startswith('mr_level')]
    cost_cols    = [c for c in df.columns if 'mr_cost' in c.lower()]
    warning_cols = [c for c in df.columns if 'warning' in c.lower()]
    company_cols = [c for c in df.columns if 'company' in c.lower()]

    if not level_cols or not cost_cols:
        return messagebox.showerror("Missing MR cols", "Need MR_level* and MR_cost* cols")
    if not company_cols:
        return messagebox.showerror("Missing Company cols", "No company columns found")

    view = view_var.get()
    html = build_map(df, level_cols, cost_cols, warning_cols,
                     company_cols, view, markup_path=markup_file)
    webbrowser.open(html)
    messagebox.showinfo("Done", f"Map saved to:\n{html}")

if __name__=="__main__":
    root = tk.Tk()
    root.title("MR / Utility Web Map")
    root.geometry("360x200")

    tk.Label(root, text="Select View:").pack(anchor='w', padx=10, pady=(10,0))
    view_var = tk.StringVar(value='MR')
    tk.Radiobutton(root, text="MR View", variable=view_var, value='MR').pack(anchor='w', padx=20)
    tk.Radiobutton(root, text="Utility View", variable=view_var, value='Utility').pack(anchor='w', padx=20)

    tk.Button(root, text="Load Markup (GeoJSON)", command=on_load_markup)\
      .pack(anchor='w', fill='x', padx=10, pady=5)
    tk.Button(root, text="Load XLSX & Generate Map", command=on_generate)\
      .pack(expand=True, fill='both', padx=10, pady=10)

    root.mainloop()