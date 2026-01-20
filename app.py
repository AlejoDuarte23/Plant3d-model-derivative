import base64
import viktor as vkt
from typing import Any
from aps_viewer_sdk import APSViewer
from aps_viewer_sdk.helper import get_all_model_properties, get_metadata_viewables
import plotly.graph_objects as go
from plotly.subplots import make_subplots


def to_md_urn(value: str) -> str:
    """Convert URN to base64 encoded format for Model Derivative API calls."""
    if value.startswith("urn:"):
        encoded = base64.urlsafe_b64encode(value.encode("utf-8")).decode("utf-8")
        return encoded.rstrip("=")
    return value.rstrip("=")


def find_prop_any_group(obj_props: dict[str, Any], key: str) -> Any | None:
    """
    Properties come as:
      obj["properties"] = { "Group A": {"Prop1": ...}, "Group B": {"Tag": ...}, ... }
    This searches every group for the given key.
    """
    if not isinstance(obj_props, dict):
        return None

    for group_props in obj_props.values():
        if isinstance(group_props, dict) and key in group_props:
            return group_props.get(key)
    return None


def build_tag_index(
    properties_payload: dict[str, Any],
    *,
    pid_keys: tuple[str, ...] = ("PnPID", "PId", "PID", "P&ID"),
    tag_key: str = "Tag",
) -> dict[str, dict[str, Any]]:
    """
    Returns:
      {
        "AV-309": {
           "objectid": 123,
           "name": "ACPPASSET [4612B]",
           "pid": 714,
           "properties": {... original grouped properties ...}
        },
        ...
      }

    If Tag repeats, it keeps the first and appends a suffix for the rest:
      "AV-309#123", "AV-309#456", ...
    """
    data = properties_payload.get("data", {})
    collection = data.get("collection", [])
    if not isinstance(collection, list):
        return {}

    out: dict[str, dict[str, Any]] = {}

    for obj in collection:
        if not isinstance(obj, dict):
            continue

        obj_props = obj.get("properties")
        if not isinstance(obj_props, dict):
            continue

        tag_val = find_prop_any_group(obj_props, tag_key)
        if tag_val is None:
            continue

        tag = str(tag_val).strip()
        if not tag:
            continue

        pid = None
        for k in pid_keys:
            pid = find_prop_any_group(obj_props, k)
            if pid is not None:
                break
        if pid is None:
            continue  # only keep items that have a P&ID id

        objectid = obj.get("objectid")
        name = obj.get("name")

        record = {
            "objectid": objectid,
            "name": name,
            "pid": pid,
            "properties": obj_props,
        }

        # Handle duplicate tags safely
        if tag not in out:
            out[tag] = record
        else:
            suffix = f"#{objectid}" if objectid is not None else "#dup"
            out[f"{tag}{suffix}"] = record

    return out


def build_class_name_counts(properties_payload: dict[str, Any]) -> dict[str, int]:
    """
    Count PID elements grouped by Class Name.
    
    Returns:
      {
        "Double_Seat_4_Port": 5,
        "Single_Valve": 12,
        ...
      }
    """
    data = properties_payload.get("data", {})
    collection = data.get("collection", [])
    if not isinstance(collection, list):
        return {}

    counts: dict[str, int] = {}

    for obj in collection:
        if not isinstance(obj, dict):
            continue

        obj_props = obj.get("properties")
        if not isinstance(obj_props, dict):
            continue

        # Only count elements that have a PID (are part of the P&ID)
        pid = None
        for k in ("PnPID", "PId", "PID", "P&ID"):
            pid = find_prop_any_group(obj_props, k)
            if pid is not None:
                break
        if pid is None:
            continue

        # Get Class Name
        class_name = find_prop_any_group(obj_props, "Class Name")
        if class_name is None:
            class_name = "Unknown"
        else:
            class_name = str(class_name).strip()
            if not class_name:
                class_name = "Unknown"

        counts[class_name] = counts.get(class_name, 0) + 1

    return counts


@vkt.memoize
def get_class_name_counts_cached(*, token: str, urn_bs64: str, model_guid: str) -> dict[str, int]:
    """
    Cached function to get all model properties and build Class Name counts.
    This is memoized to avoid repeated API calls for the same model.
    """
    properties_payload = get_all_model_properties(
        token=token,
        urn_bs64=urn_bs64,
        model_guid=model_guid
    )
    
    return build_class_name_counts(properties_payload)


@vkt.memoize
def get_metadata_views_cached(*, token: str, urn_bs64: str) -> list[dict[str, Any]]:
    """
    Cached function to get metadata viewables.
    This is memoized to avoid repeated API calls for the same URN.
    """
    metadata_views = get_metadata_viewables(token, urn_bs64)
    return metadata_views if metadata_views else []


@vkt.memoize
def get_tag_index_cached(*, token: str, urn_bs64: str, model_guid: str) -> dict[str, dict[str, Any]]:
    """
    Cached function to get all model properties and build Tag index.
    This is memoized to avoid repeated API calls for the same model.
    """
    properties_payload = get_all_model_properties(
        token=token,
        urn_bs64=urn_bs64,
        model_guid=model_guid
    )
    
    tag_index = build_tag_index(properties_payload)
    return tag_index


def get_viewables(params, **kwargs):
    """Gets option list elements name - metadata guid for properties"""
    autodesk_file = params.autodesk_file
    if not autodesk_file:
        return []

    integration = vkt.external.OAuth2Integration("aps-integration-viktor")
    token = integration.get_access_token()
    version = autodesk_file.get_latest_version(token)
    version_urn = version.urn
    urn_bs64 = to_md_urn(version_urn)
    
    # Get cached metadata viewables (memoized to avoid repeated API calls)
    metadata_views = get_metadata_views_cached(token=token, urn_bs64=urn_bs64)
    
    if not metadata_views:
        return []
    
    # Create OptionListElements with name as label and metadata guid as value
    options = []
    for viewable in metadata_views:
        name = viewable.get("name", "Unknown View")
        guid = viewable.get("guid")
        role = viewable.get("role", "")
        if guid:
            label = f"{name} ({role})" if role else name
            options.append(vkt.OptionListElement(label=label, value=guid))
    
    return options


def get_tag_options(params, **kwargs):
    """Gets option list elements for PID tags from the selected view"""
    autodesk_file = params.autodesk_file
    if not autodesk_file:
        return []
    
    selected_guid = params.selected_view
    if not selected_guid:
        return []

    integration = vkt.external.OAuth2Integration("aps-integration-viktor")
    token = integration.get_access_token()
    version = autodesk_file.get_latest_version(token)
    version_urn = version.urn
    urn_bs64 = to_md_urn(version_urn)
    
    # Get cached tag index
    tag_index = get_tag_index_cached(
        token=token,
        urn_bs64=urn_bs64,
        model_guid=selected_guid
    )
    
    if not tag_index:
        return []
    
    options = []
    for tag in sorted(tag_index.keys()):
        options.append(vkt.OptionListElement(label=tag, value=tag))
    
    return options


class Parametrization(vkt.Parametrization):
    title = vkt.Text("""# Plant 3D - Integration
This application allows users to view Plant 3D models and explore PID tags and their properties.
Select a Plant 3D model, choose a viewable, and configure tag parameters.
    """)
    
    header1 = vkt.Text("""## 1. Select Plant 3D Model""")
    autodesk_file = vkt.AutodeskFileField(
        "Plant 3D Field",
        oauth2_integration="aps-integration-viktor"
    )
    
    lbk0 = vkt.LineBreak()
    header2 = vkt.Text("""## 2. Select Viewable""")
    selected_view = vkt.OptionField("Select Plant3D Viewable", options=get_viewables)
    
    lbk1 = vkt.LineBreak()
    header3 = vkt.Text("""## 3. Configure Tag Parameters""")
    tag_params = vkt.DynamicArray("Tag Parameters", row_label="Tag", copylast=True)
    tag_params.tag = vkt.OptionField("PID Tag", options=get_tag_options)
    tag_params.param_name = vkt.TextField("Parameter Name")
    tag_params.value = vkt.TextField("Value")


class Controller(vkt.Controller):
    parametrization = Parametrization

    @vkt.WebView("Plant3D View", duration_guess=30)
    def dwg_view(self, params, **kwargs) -> vkt.WebResult:
        autodesk_file = params.autodesk_file
        if not autodesk_file:
            return vkt.WebResult(html="<p>Please select a Plant 3D file.</p>")

        integration = vkt.external.OAuth2Integration("aps-integration-viktor")
        token = integration.get_access_token()
        version = autodesk_file.get_latest_version(token)
        version_urn = version.urn
        viewer = APSViewer(urn=version_urn, token=token)
        html = viewer.write()
        return vkt.WebResult(html=html)
    
    @vkt.PlotlyView("Quantity Takeoff", duration_guess=30)
    def quantity_takeoff_view(self, params, **kwargs) -> vkt.PlotlyResult:
        """Combined pie chart and bar chart showing PID elements grouped by Class Name."""
        autodesk_file = params.autodesk_file
        selected_guid = params.selected_view
        
        if not autodesk_file or not selected_guid:
            fig = go.Figure()
            fig.add_annotation(
                text="Please select a Plant 3D file and viewable",
                xref="paper", yref="paper",
                x=0.5, y=0.5, showarrow=False,
                font=dict(size=16)
            )
            fig.update_layout(
                title="Quantity Takeoff - PID Elements by Class Name",
                showlegend=False
            )
            return vkt.PlotlyResult(fig)

        integration = vkt.external.OAuth2Integration("aps-integration-viktor")
        token = integration.get_access_token()
        version = autodesk_file.get_latest_version(token)
        version_urn = version.urn
        urn_bs64 = to_md_urn(version_urn)
        
        # Get cached class name counts
        class_counts = get_class_name_counts_cached(
            token=token,
            urn_bs64=urn_bs64,
            model_guid=selected_guid
        )
        
        if not class_counts:
            fig = go.Figure()
            fig.add_annotation(
                text="No PID elements found in the selected view",
                xref="paper", yref="paper",
                x=0.5, y=0.5, showarrow=False,
                font=dict(size=16)
            )
            fig.update_layout(
                title="Quantity Takeoff - PID Elements by Class Name",
                showlegend=False
            )
            return vkt.PlotlyResult(fig)
        
        # Sort by count descending
        sorted_items = sorted(class_counts.items(), key=lambda x: x[1], reverse=True)
        # Format labels: replace underscores with spaces for readability
        labels = [item[0].replace('_', ' ') for item in sorted_items]
        values = [item[1] for item in sorted_items]
        total_count = sum(values)
        
        # Create subplots: pie chart on left, horizontal bar chart on right
        fig = make_subplots(
            rows=1, cols=2,
            specs=[[{"type": "pie"}, {"type": "bar"}]],
            column_widths=[0.35, 0.65],
            horizontal_spacing=0.35,
            subplot_titles=("Distribution Overview", "Count by Class Name")
        )
        
        # Pie chart - show top 8 categories, group rest as "Other"
        top_n = 8
        if len(labels) > top_n:
            pie_labels = labels[:top_n] + ["Other"]
            pie_values = values[:top_n] + [sum(values[top_n:])]
        else:
            pie_labels = labels
            pie_values = values
        
        # Color palette for consistency
        colors = [
            '#636EFA', '#EF553B', '#00CC96', '#AB63FA', '#FFA15A',
            '#19D3F3', '#FF6692', '#B6E880', '#FF97FF', '#FECB52'
        ]
        
        fig.add_trace(
            go.Pie(
                labels=pie_labels,
                values=pie_values,
                hole=0.4,
                textinfo='percent',
                textposition='outside',
                marker=dict(colors=colors[:len(pie_labels)]),
                hovertemplate="<b>%{label}</b><br>Count: %{value}<br>Percentage: %{percent}<extra></extra>",
                showlegend=False
            ),
            row=1, col=1
        )
        
        # Horizontal bar chart - all categories, sorted ascending for bottom-to-top reading
        bar_labels = labels[::-1]  # Reverse so highest is at top
        bar_values = values[::-1]
        
        # Assign colors based on value (gradient effect)
        max_val = max(bar_values) if bar_values else 1
        bar_colors = [f'rgba(99, 110, 250, {0.4 + 0.6 * (v / max_val)})' for v in bar_values]
        
        fig.add_trace(
            go.Bar(
                x=bar_values,
                y=bar_labels,
                orientation='h',
                marker=dict(
                    color=bar_colors,
                    line=dict(color='rgba(99, 110, 250, 1)', width=1)
                ),
                text=bar_values,
                textposition='outside',
                hovertemplate="<b>%{y}</b><br>Count: %{x}<br>Percentage: %{customdata:.1f}%<extra></extra>",
                customdata=[v / total_count * 100 for v in bar_values],
                showlegend=False
            ),
            row=1, col=2
        )
        
        # Calculate dynamic height based on number of categories
        chart_height = max(500, len(labels) * 25 + 150)
        
        fig.update_layout(
            title=dict(
                text=f"<b>Quantity Takeoff - PID Elements by Class Name</b><br><sup>Total: {total_count} elements | {len(labels)} unique classes</sup>",
                x=0.5,
                xanchor='center',
                font=dict(size=18)
            ),
            height=chart_height,
            margin=dict(t=100, b=50, l=200, r=80),
            showlegend=False
        )
        
        # Update bar chart x-axis
        fig.update_xaxes(
            title_text="Count",
            row=1, col=2,
            gridcolor='lightgray',
            range=[0, max(values) * 1.15]  # Add space for labels
        )
        
        # Update bar chart y-axis
        fig.update_yaxes(
            title_text="",
            row=1, col=2,
            tickfont=dict(size=10)
        )
        
        return vkt.PlotlyResult(fig)
    
    def build_tag_properties_dict(self, params, **kwargs) -> dict[str, Any]:
        """Convert dynamic array into structured properties dictionary"""
        tag_params = params.tag_params
        if not tag_params:
            return {"version": 1, "items": []}
        
        # Group parameters by tag
        tag_groups: dict[str, dict[str, str]] = {}
        for row in tag_params:
            tag = row.get("tag")
            param_name = row.get("param_name")
            value = row.get("value")
            
            # Skip incomplete rows
            if not tag or not param_name or not value:
                continue
            
            # Initialize tag group if not exists
            if tag not in tag_groups:
                tag_groups[tag] = {}
            
            # Add parameter to the tag's properties
            tag_groups[tag][param_name] = value
        
        # Build the final structure
        items = []
        for tag, properties in tag_groups.items():
            items.append({
                "match": {"tag": tag},
                "properties": properties
            })
        
        return {
            "version": 1,
            "items": items
        }
