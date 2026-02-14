"""
Worm gear pair analysis model.

Computes contact stress, root stress, torque, efficiency and fatigue
across one full mesh cycle (phi = 0..2*pi).

This is a lightweight proxy model for trend analysis, not a full FEM solver.
"""

import math
import numpy as np


def _interp_Et(wheel, temp_C):
    """Interpolate elastic modulus E(T) from material card."""
    pts = wheel.get("elastic_T", {}).get("points_C_GPa", [])
    if not pts:
        return wheel.get("elastic", {}).get("E_GPa", 2.0)
    pts = sorted(pts, key=lambda p: p[0])
    temps = [p[0] for p in pts]
    vals = [p[1] for p in pts]
    return float(np.interp(temp_C, temps, vals))


def _interp_sn(sn_list, N):
    """Interpolate allowable stress from S-N curve [[N, MPa], ...]."""
    if not sn_list:
        return None
    pts = sorted(sn_list, key=lambda p: p[0])
    ns = [math.log10(p[0]) for p in pts]
    ss = [p[1] for p in pts]
    logN = math.log10(N)
    return float(np.interp(logN, ns, ss))


def compute_worm_cycle(inp, steel, wheel):
    """
    Main computation entry point.

    Parameters
    ----------
    inp : dict
        All input parameters as string values.
    steel : dict
        Worm material JSON data.
    wheel : dict
        Wheel material JSON data.

    Returns
    -------
    dict with keys:
        phi, p_contact_MPa, sigma_root_MPa, T2_Nm, eta, Nc_proxy, meta
    """
    # Parse inputs
    T1 = float(inp.get("T1_Nm", 6.0))
    n1 = float(inp.get("n1_rpm", 3000))
    ratio = float(inp.get("ratio", 25))
    z1 = int(float(inp.get("z1", 2)))
    mn = float(inp.get("mn_mm", 2.5))
    q = float(inp.get("q", 10))
    x1 = float(inp.get("x1", 0.0))
    x2 = float(inp.get("x2", 0.0))
    a_target_txt = inp.get("a_target_mm", "").strip()
    b = float(inp.get("b_mm", 18))
    alpha_n = math.radians(float(inp.get("alpha_n_deg", 20)))
    mu = float(inp.get("mu", 0.06))
    KA = float(inp.get("KA", 1.1))
    KV = float(inp.get("KV", 1.05))
    KHb = float(inp.get("KHb", 1.0))
    KFb = float(inp.get("KFb", 1.0))
    temp_C = float(inp.get("temp_C", 80))
    life_h = float(inp.get("life_h", 3000))
    steps = int(float(inp.get("steps", 720)))
    rho_f = float(inp.get("rho_f_mm", 0.6))

    # Worm geometry
    z2 = int(round(ratio * z1))
    d1 = (q + 2.0 * x1) * mn          # worm pitch diameter
    da1 = d1 + 2.0 * mn                # worm tip diameter
    df1 = d1 - 2.4 * mn                # worm root diameter
    d2 = (z2 + 2.0 * x2) * mn          # wheel pitch diameter
    da2 = d2 + 2.0 * mn * (1.0 + x2)   # wheel tip diameter
    df2 = d2 - 2.0 * mn * (1.2 - x2)   # wheel root diameter
    a_calc = 0.5 * (d1 + d2)

    a_target = float(a_target_txt) if a_target_txt else None
    a_mm = a_target if a_target is not None else a_calc
    delta_a = (a_target - a_calc) if a_target is not None else None

    # Lead angle
    gamma = math.atan2(z1 * mn, d1) if d1 > 0 else math.radians(5)
    gamma_deg = math.degrees(gamma)

    # Material properties
    E1 = steel.get("elastic", {}).get("E_GPa", 210.0)
    nu1 = steel.get("elastic", {}).get("nu", 0.3)
    E2 = _interp_Et(wheel, temp_C)
    nu2 = wheel.get("nu", 0.4)

    # Equivalent elastic modulus (Hertz)
    Eprime = 2.0 / ((1 - nu1**2) / E1 + (1 - nu2**2) / E2)  # GPa

    # Efficiency (simplified friction model)
    eta0 = math.cos(alpha_n) - mu / math.tan(gamma)
    eta0 /= (math.cos(alpha_n) + mu * math.tan(gamma)) if math.cos(alpha_n) + mu * math.tan(gamma) != 0 else 1
    eta0 = max(0.3, min(eta0, 0.98))

    # Output torque
    T2_base = T1 * ratio * eta0  # N*m

    # Phase sweep
    phi = np.linspace(0, 2 * np.pi, steps, endpoint=False)

    # Mesh stiffness variation proxy (sinusoidal)
    Nc_proxy = 1.0 + 0.15 * np.sin(z1 * phi) + 0.08 * np.cos(2 * z1 * phi)
    eta = eta0 * (1.0 - 0.015 * (1.0 - np.cos(z1 * phi)))

    # T2 ripple
    T2_Nm = T2_base * (1.0 + 0.04 * np.sin(z1 * phi) + 0.02 * np.sin(2 * z1 * phi))

    # Contact stress proxy (Hertz-like)
    # p ~ sqrt(Fn * E' / (rho * b))
    Fn_base = T1 * 1000.0 / (0.5 * d1) if d1 > 0 else 0  # N (T in Nm, d in mm)
    rho_eq = 0.5 * d1 * math.sin(alpha_n) if d1 > 0 else 1.0  # mm equivalent curvature

    K_total_H = KA * KV * KHb
    p_base = 0.418 * math.sqrt(Fn_base * K_total_H * Eprime * 1000 / (rho_eq * b)) if (rho_eq * b) > 0 else 0
    # Phase variation
    p_contact_MPa = p_base * (1.0 + 0.06 * np.sin(z1 * phi) + 0.03 * np.cos(2 * z1 * phi))

    # Root stress proxy (Lewis-like)
    # sigma_F ~ Fn / (b * mn * Y)
    Y_F = 2.2  # form factor proxy
    K_total_F = KA * KV * KFb
    sigma_base = (Fn_base * K_total_F * Y_F) / (b * mn) if (b * mn) > 0 else 0
    sigma_root_MPa = sigma_base * (1.0 + 0.08 * np.sin(z1 * phi) + 0.04 * np.cos(2 * z1 * phi))

    # S-N lookup for safety factors
    sn = wheel.get("SN", {})
    contact_sn = sn.get("contact_allow_MPa_vs_N", [])
    root_sn = sn.get("root_allow_MPa_vs_N", [])

    N_life = n1 * 60 * life_h / ratio  # wheel cycles

    sigma_root_max = float(np.max(sigma_root_MPa))
    p_contact_max = float(np.max(p_contact_MPa))

    SF_root = None
    if root_sn:
        allow_root = _interp_sn(root_sn, max(N_life, 1))
        if allow_root and sigma_root_max > 0:
            SF_root = allow_root / sigma_root_max

    SF_contact = None
    if contact_sn:
        allow_contact = _interp_sn(contact_sn, max(N_life, 1))
        if allow_contact and p_contact_max > 0:
            SF_contact = allow_contact / p_contact_max

    # Miner damage proxy (rainflow simplified: assume sinusoidal load)
    # For a sinusoidal stress, equivalent amplitude = (max-min)/2
    sigma_amp = 0.5 * (float(np.max(sigma_root_MPa)) - float(np.min(sigma_root_MPa)))
    damage_root = 0.0
    if root_sn and sigma_amp > 0:
        # Simplified: each revolution contributes z1 cycles of sigma_amp
        n_revs = n1 * 60 * life_h
        n_cycles_total = n_revs * z1
        # Find allowable N for this amplitude from SN curve (inverse lookup)
        pts = sorted(root_sn, key=lambda p: p[1], reverse=True)
        log_ns = [math.log10(p[0]) for p in pts]
        stresses = [p[1] for p in pts]
        if sigma_amp >= min(stresses):
            logN_allow = float(np.interp(sigma_amp, stresses[::-1], log_ns[::-1]))
            N_allow = 10 ** logN_allow
            damage_root = n_cycles_total / N_allow if N_allow > 0 else 999.0

    meta = {
        "z1": z1,
        "z2": z2,
        "d1_mm": d1,
        "da1_mm": da1,
        "df1_mm": df1,
        "d2_mm": d2,
        "da2_mm": da2,
        "df2_mm": df2,
        "a_mm": a_mm,
        "a_calc_mm": a_calc,
        "a_target_mm": a_target,
        "delta_a_mm": delta_a,
        "x1": x1,
        "x2": x2,
        "gamma_deg": gamma_deg,
        "eta0": eta0,
        "Eprime_GPa": Eprime,
        "KA": KA,
        "KV": KV,
        "KHb": KHb,
        "KFb": KFb,
        "SF_root": SF_root,
        "SF_contact": SF_contact,
        "damage_root": damage_root,
        "N_life": N_life,
        "Fn_base_N": Fn_base,
        "rho_eq_mm": rho_eq,
    }

    return {
        "phi": phi,
        "p_contact_MPa": p_contact_MPa,
        "sigma_root_MPa": sigma_root_MPa,
        "T2_Nm": T2_Nm,
        "eta": eta,
        "Nc_proxy": Nc_proxy,
        "meta": meta,
    }
