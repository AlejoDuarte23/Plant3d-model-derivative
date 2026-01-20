import base64
import viktor as vkt
from typing import Any
from aps_viewer_sdk import APSViewer
from aps_viewer_sdk.helper import get_all_model_properties, get_metadata_viewables
from aps_viewer_sdk.client import ElementsInScene
import plotly.graph_objects as go
from plotly.subplots import make_subplots


# Hardcoded property options based on Plant 3D element properties
PROPERTY_OPTIONS = [
    "Description",
    "Status", 
    "Tag",
    "Size",
    "Spec",
    "Service",
    "Insulation Thickness",
    "Type",
    "Number",
    "Capacity",
    "PnPID",
    "PnPGuid",
    "Class Name",
    "Manufacturer",
]


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
def get_properties_payload_cached(*, token: str, urn_bs64: str, model_guid: str) -> dict[str, Any]:
    """
    Cached function to get all model properties payload.
    This is memoized to avoid repeated API calls for the same model.
    """
    return get_all_model_properties(
        token=token,
        urn_bs64=urn_bs64,
        model_guid=model_guid
    )


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


def get_class_name_options(params, **kwargs):
    """Gets option list elements for Class Names from the selected view"""
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
    
    # Get cached class name counts
    class_counts = get_class_name_counts_cached(
        token=token,
        urn_bs64=urn_bs64,
        model_guid=selected_guid
    )
    
    if not class_counts:
        return []
    
    # Sort by count descending and format nicely
    sorted_items = sorted(class_counts.items(), key=lambda x: x[1], reverse=True)
    options = []
    for class_name, count in sorted_items:
        display_name = class_name.replace('_', ' ')
        options.append(vkt.OptionListElement(
            label=f"{display_name} ({count})", 
            value=class_name
        ))
    
    return options


class Parametrization(vkt.Parametrization):
    title = vkt.Text("""# Plant 3D - Integration
This application allows users to view Plant 3D models, explore PID tags, and perform QA/QC checks.
Select a Plant 3D model, choose a viewable, and configure filters to highlight elements.
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
    header3 = vkt.Text("""## 3. QA/QC Filter Configuration
Configure filters to highlight elements in the QA/QC view. Each filter row will be highlighted with the selected color.
Select multiple Class Names per row if needed.
    """)
    qc_filters = vkt.DynamicArray("QA/QC Filters", row_label="Filter", copylast=True)
    qc_filters.class_names = vkt.MultiSelectField("Class Names", options=get_class_name_options)
    qc_filters.property_name = vkt.OptionField("Property", options=PROPERTY_OPTIONS)
    qc_filters.expected_value = vkt.TextField("Expected Value (leave empty to match any)")
    qc_filters.highlight_color = vkt.ColorField("Highlight Color", default=vkt.Color.blue())


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

    
    @vkt.WebView("QA/QC View", duration_guess=30)
    def qaqc_view(self, params, **kwargs) -> vkt.WebResult:
        """QA/QC view that highlights elements based on filter criteria."""
        autodesk_file = params.autodesk_file
        selected_guid = params.selected_view
        
        if not autodesk_file:
            return vkt.WebResult(html="<p>Please select a Plant 3D file.</p>")
        
        if not selected_guid:
            return vkt.WebResult(html="<p>Please select a viewable.</p>")

        integration = vkt.external.OAuth2Integration("aps-integration-viktor")
        token = integration.get_access_token()
        version = autodesk_file.get_latest_version(token)
        version_urn = version.urn
        urn_bs64 = to_md_urn(version_urn)
        
        # Create viewer
        viewer = APSViewer(urn=version_urn, token=token)
        
        # Get viewables and set the selected view
        viewables = viewer.get_viewables(urn_bs64)
        if viewables:
            selected_viewable = next(
                (v for v in viewables if v.get("guid") == selected_guid),
                viewables[0]
            )
            viewer.set_view_guid(
                selected_viewable["guid"],
                selected_viewable.get("name", "View"),
                selected_viewable.get("role", "3d")
            )
        
        # Parse filters from params
        qc_filters = params.qc_filters or []
        
        # If no filters, show viewer without highlighting
        if not qc_filters or all(not f.get("class_names") for f in qc_filters):
            html = viewer.write()
            return vkt.WebResult(html=html)
        
        # Get properties payload
        properties_payload = get_properties_payload_cached(
            token=token,
            urn_bs64=urn_bs64,
            model_guid=selected_guid
        )
        
        data = properties_payload.get("data", {})
        collection = data.get("collection", [])
        
        # Default color if none specified
        default_color = "#FF0000"
        
        # Build filter criteria with colors from user selection
        filter_criteria = []
        for idx, f in enumerate(qc_filters):
            class_names = f.get("class_names") or []
            if not class_names:
                continue
            
            # Get color from user selection, use default if not set
            user_color = f.get("highlight_color")
            if user_color:
                color_hex = user_color.hex
            else:
                color_hex = default_color
            
            filter_criteria.append({
                "class_names": class_names,
                "property_name": f.get("property_name"),
                "expected_value": (f.get("expected_value") or "").strip(),
                "color": color_hex
            })
        
        if not filter_criteria:
            html = viewer.write()
            return vkt.WebResult(html=html)
        
        # Find matching elements
        highlight_elements: list[ElementsInScene] = []
        
        for obj in collection:
            if not isinstance(obj, dict):
                continue
            
            obj_props = obj.get("properties")
            if not isinstance(obj_props, dict):
                continue
            
            external_id = obj.get("externalId")
            if not external_id:
                continue
            
            # Get element's class name
            obj_class_name = find_prop_any_group(obj_props, "Class Name")
            if obj_class_name is None:
                continue
            
            # Check if element matches any filter
            for criteria in filter_criteria:
                # Check if class name is in the selected class names
                if obj_class_name not in criteria["class_names"]:
                    continue
                
                # If property name is specified, check property match
                prop_name = criteria.get("property_name")
                expected_val = criteria.get("expected_value")
                
                if prop_name and expected_val:
                    # Check if property matches expected value
                    actual_val = find_prop_any_group(obj_props, prop_name)
                    if actual_val is None:
                        continue
                    if str(actual_val).strip().lower() != expected_val.lower():
                        continue
                elif prop_name and not expected_val:
                    # Property specified but no value - just check if property exists
                    actual_val = find_prop_any_group(obj_props, prop_name)
                    if actual_val is None:
                        continue
                
                # Element matches - highlight it with this filter's color
                highlight_elements.append({
                    "externalElementId": external_id,
                    "color": criteria["color"]
                })
                break  # Element matched, use first matching filter's color
        
        # Apply highlighting
        if highlight_elements:
            viewer.highlight_elements(highlight_elements)
        
        html = viewer.write()
        return vkt.WebResult(html=html)
    