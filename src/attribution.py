"""
src/attribution.py - Multi-Touch Attribution Engine
Models First-Touch, Last-Touch, and Linear attribution across customer journeys.
Enforces credit conservation: sum(credits) == total conversions.
"""

def compute_attribution_models(touchpoint_data, channel_map: dict = None) -> dict:
    """
    Computes First-Touch, Last-Touch, and Linear conversion credits by customer journey.
    """
    records = touchpoint_data.to_dict(orient="records") if hasattr(touchpoint_data, "to_dict") else touchpoint_data
    
    # Group touchpoints by customer_id ordered by touchpoint_timestamp or touchpoint_id
    journeys = {}
    for r in records:
        cust_id = r.get("customer_id")
        if cust_id not in journeys:
            journeys[cust_id] = []
        journeys[cust_id].append(r)
        
    for cust_id in journeys:
        journeys[cust_id].sort(key=lambda x: str(x.get("touchpoint_timestamp", x.get("touchpoint_id", 0))))

    # Default channel id -> channel name mapping if not provided
    default_names = {
        1: "Google Search Ads",
        2: "Meta Performance Ads",
        3: "Email Marketing",
        4: "Influencer Partnerships",
        5: "Organic Search"
    }
    if channel_map is None:
        channel_map = default_names

    channels = list(channel_map.values())
    first_touch = {ch: 0.0 for ch in channels}
    last_touch = {ch: 0.0 for ch in channels}
    linear = {ch: 0.0 for ch in channels}
    
    total_conversions = 0.0

    for cust_id, path in journeys.items():
        if not path:
            continue
        # Check if customer converted
        converted = any(r.get("is_conversion") == 1 or r.get("converted") == 1 for r in path)
        if not converted:
            continue
            
        total_conversions += 1.0
        n_touch = len(path)
        
        # First touch channel
        ch_first_id = path[0].get("channel_id")
        ch_first = channel_map.get(ch_first_id, f"Channel {ch_first_id}")
        first_touch[ch_first] = first_touch.get(ch_first, 0.0) + 1.0
        
        # Last touch channel
        ch_last_id = path[-1].get("channel_id")
        ch_last = channel_map.get(ch_last_id, f"Channel {ch_last_id}")
        last_touch[ch_last] = last_touch.get(ch_last, 0.0) + 1.0
        
        # Linear credit (1 / N per touchpoint)
        credit_per_touch = 1.0 / n_touch
        for step in path:
            ch_id = step.get("channel_id")
            ch_name = channel_map.get(ch_id, f"Channel {ch_id}")
            linear[ch_name] = linear.get(ch_name, 0.0) + credit_per_touch

    # Round credits
    first_touch = {k: round(v, 2) for k, v in first_touch.items()}
    last_touch = {k: round(v, 2) for k, v in last_touch.items()}
    linear = {k: round(v, 2) for k, v in linear.items()}

    return {
        "total_conversions": total_conversions,
        "first_touch": first_touch,
        "last_touch": last_touch,
        "linear": linear
    }
