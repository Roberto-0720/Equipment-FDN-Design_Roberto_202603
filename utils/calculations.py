"""
Core calculation engine for Octagonal Spread Footing Analysis.
Replicates all formulas from Sheet 1 "Octagonal Footing" of EQ.xlsm.
Reference: PIP STE03350 - "Vertical Vessel Foundation Design Guide" (2007)
"""

import math
from data.tables import CORNERS_TABLE, FLAT_SIDES_TABLE, interpolate_K_L


class OctagonFootingCalc:
    """
    Octagonal Spread Footing Analysis Calculator.
    For Assumed Rigid Footing Base with Octagonal Pier
    Supporting a Vertical Round Tank, Vessel, or Stack.
    """

    def __init__(self, Df, Tf, Dp, hp, gamma_c, Ds, gamma_s, Kp, mu, Q, q_allow):
        """
        Initialize with footing input parameters.
        
        Args:
            Df: Ftg. Base Length (flat-to-flat dimension), m
            Tf: Ftg. Base Thickness, m
            Dp: Oct. Pier Length, m
            hp: Oct. Pier Height, m
            gamma_c: Concrete Unit Wt., kN/m³
            Ds: Soil Depth, m
            gamma_s: Soil Unit Wt., kN/m³
            Kp: Pass. Press. Coef.
            mu: Coef. of Base Friction
            Q: Uniform Surcharge, kN/m²
            q_allow: SB Capacity, kN/m²
        """
        self.Df = Df
        self.Tf = Tf
        self.Dp = Dp
        self.hp = hp
        self.gamma_c = gamma_c
        self.Ds = Ds
        self.gamma_s = gamma_s
        self.Kp = Kp
        self.mu = mu
        self.Q = Q
        self.q_allow = q_allow

        # Pre-compute geometry
        self._compute_footing_properties()
        self._compute_pier_properties()
        self._compute_weights()

    def _compute_footing_properties(self):
        """Compute Footing Base Properties."""
        Df = self.Df
        # Bf = Cf * sin(45°) = 0.2928932 * Df
        self.Cf = math.tan(math.radians(22.5)) * Df
        self.Bf = self.Cf * math.sin(math.radians(45))
        # Ef = Df / cos(22.5°) = 1.082392 * Df
        self.Ef = Df / math.cos(math.radians(22.5))
        # Af = (1 - 2*(tan(22.5)*sin(45))²) * Df² = 0.8284272 * Df²
        self.Af = (1 - 2 * (math.tan(math.radians(22.5)) * math.sin(math.radians(45))) ** 2) * Df ** 2
        # Vf = Af * Tf
        self.Vf = self.Af * self.Tf
        # If = 0.054738 * Df⁴
        self.If = 0.054738 * Df ** 4

    def _compute_pier_properties(self):
        """Compute Pier Properties."""
        Dp = self.Dp
        self.Cp = math.tan(math.radians(22.5)) * Dp
        self.Bp = self.Cp * math.sin(math.radians(45))
        self.Ep = Dp / math.cos(math.radians(22.5))
        self.Ap = (1 - 2 * (math.tan(math.radians(22.5)) * math.sin(math.radians(45))) ** 2) * Dp ** 2
        self.Vp = self.Ap * self.hp

    def _compute_weights(self):
        """Compute Pier, Surcharge, Soil, and Footing Base Weights."""
        # Wp = Vp * γc
        self.Wp = self.Vp * self.gamma_c
        # Wq = (Af - Ap) * Q
        self.Wq = (self.Af - self.Ap) * self.Q
        # Ws = (Af - Ap) * Ds * γs
        self.Ws = (self.Af - self.Ap) * self.Ds * self.gamma_s
        # Wf = Vf * γc
        self.Wf = self.Vf * self.gamma_c

    def calculate(self, P, H, M):
        """
        Run full analysis for a given load combination.
        
        Args:
            P: Applied Vert. Load, kN
            H: Applied Horiz. Load, kN
            M: Applied Moment, kN-m
            
        Returns:
            dict with all calculation results
        """
        results = {}
        
        # ---- Input Data ----
        results['Df'] = self.Df
        results['Tf'] = self.Tf
        results['Dp'] = self.Dp
        results['hp'] = self.hp
        results['gamma_c'] = self.gamma_c
        results['Ds'] = self.Ds
        results['gamma_s'] = self.gamma_s
        results['Kp'] = self.Kp
        results['mu'] = self.mu
        results['Q'] = self.Q
        results['q_allow'] = self.q_allow
        results['P'] = P
        results['H'] = H
        results['M'] = M

        # ---- Footing Base Properties ----
        results['Bf'] = self.Bf
        results['Cf'] = self.Cf
        results['Ef'] = self.Ef
        results['Af'] = self.Af
        results['Vf'] = self.Vf
        results['If'] = self.If

        # ---- Pier Properties ----
        results['Bp'] = self.Bp
        results['Cp'] = self.Cp
        results['Ep'] = self.Ep
        results['Ap'] = self.Ap
        results['Vp'] = self.Vp

        # ---- Weights ----
        results['Wp'] = self.Wp
        results['Wq'] = self.Wq
        results['Ws'] = self.Ws
        results['Wf'] = self.Wf

        # ---- Total Resultant Load and Eccentricities ----
        SP = P + self.Wp + self.Wq + self.Ws + self.Wf
        SM = M + H * (self.hp + self.Tf)
        
        if abs(SP) < 1e-10:
            e = 0.0
        else:
            e = SM / SP
        
        e_Df = round(e / self.Df, 4) if abs(self.Df) > 1e-10 else 0.0
        
        results['SP'] = SP
        results['SM'] = SM
        results['e'] = e
        results['e_Df'] = e_Df

        # ---- Overturning Check ----
        # Mr = (SP - Wq) * (Df/2)  (surcharge not included)
        Mr = (SP - self.Wq) * (self.Df / 2)
        # Mo = SP * e
        Mo = SP * e
        
        if Mo > 0:
            FS_ot = Mr / Mo
        else:
            FS_ot = "N.A."
        
        results['Mr'] = Mr
        results['Mo'] = Mo
        results['FS_ot'] = FS_ot

        # ---- Sliding Check ----
        # PR = (1/2*Kp*γs*Ds²)*Dp + (Tf*(Kp*γs*Ds + Kp*γs*(Ds+Tf))/2)*Df
        Kp = self.Kp
        gs = self.gamma_s
        Ds = self.Ds
        Df = self.Df
        Dp = self.Dp
        Tf = self.Tf
        
        PR = (0.5 * Kp * gs * Ds ** 2) * Dp + (Tf * (Kp * gs * Ds + Kp * gs * (Ds + Tf)) / 2) * Df
        # FR = (SP - Wq) * μ
        FR = (SP - self.Wq) * self.mu
        
        if H > 0:
            FS_slid = (PR + FR) / H
        else:
            FS_slid = "N.A."
        
        results['PR'] = PR
        results['FR'] = FR
        results['FS_slid'] = FS_slid

        # ---- Bearing Pressure: Axis through CORNERS ----
        corners = self._calc_bearing_corners(SP, SM, e_Df)
        results.update({f'corners_{k}': v for k, v in corners.items()})

        # ---- Bearing Pressure: Axis through FLAT SIDES ----
        flat = self._calc_bearing_flat(SP, SM, e_Df)
        results.update({f'flat_{k}': v for k, v in flat.items()})

        # ---- Summary Results ----
        # Max gross bearing pressure
        Pmax_gross_corners = corners.get('Pmax_gross', 0)
        Pmax_gross_flat = flat.get('Pmax_gross', 0)
        
        # Handle string results ("Resize!", "N.A.")
        if isinstance(Pmax_gross_corners, str):
            Pmax_gross_corners = 0
        if isinstance(Pmax_gross_flat, str):
            Pmax_gross_flat = 0
            
        if Pmax_gross_corners >= Pmax_gross_flat:
            results['Pmax_gross'] = corners.get('Pmax_gross', 0)
            results['pct_brg_area'] = corners.get('pct_brg_area', 100)
        else:
            results['Pmax_gross'] = flat.get('Pmax_gross', 0)
            results['pct_brg_area'] = flat.get('pct_brg_area', 100)
        
        results['Pmax_net'] = max(
            corners.get('Pmax_net', 0) if not isinstance(corners.get('Pmax_net', 0), str) else 0,
            flat.get('Pmax_net', 0) if not isinstance(flat.get('Pmax_net', 0), str) else 0
        )
        
        results['Pmax_gross_max'] = max(Pmax_gross_corners, Pmax_gross_flat)
        
        return results

    def _calc_bearing_corners(self, SP, SM, e_Df):
        """
        Bearing Pressure for Overturning about Axis through Corners of Octagon.
        """
        res = {}
        
        # Section Modulus: Sf = If / (Ef/2)
        Sf = self.If / (self.Ef / 2) if self.Ef > 0 else 0
        res['Sf'] = Sf
        
        # K interpolation from corners table
        threshold = 0.1221
        if e_Df > threshold:
            K, L = interpolate_K_L(e_Df, CORNERS_TABLE)
            if K is None:
                res['K'] = "Resize!"
                res['K_Df'] = "Resize!"
                res['pct_brg_area'] = "Resize!"
                res['L'] = "Resize!"
                res['Pmax_gross'] = "Resize!"
                res['Pmin_gross'] = "Resize!"
                res['Pmax_net'] = "Resize!"
                return res
        else:
            K = 0
            L = None
        
        res['K'] = K
        K_Df = K * self.Df if isinstance(K, (int, float)) else K
        res['K_Df'] = K_Df
        
        # %Brg Area calculation
        if isinstance(K_Df, (int, float)) and K_Df <= 0.65 * self.Df:
            pct_brg = self._calc_pct_brg_area_corners(K_Df)
        else:
            pct_brg = "Resize!"
        res['pct_brg_area'] = pct_brg
        
        # L coefficient
        if e_Df > threshold:
            res['L'] = L if L is not None else "Resize!"
        else:
            res['L'] = "N.A."
        
        # Gross Bearing Pressure
        if isinstance(res.get('L'), str) and res['L'] == "Resize!":
            res['Pmax_gross'] = "Resize!"
            res['Pmin_gross'] = "Resize!"
        elif e_Df <= threshold:
            # P(max) = SP/Af + SM/Sf
            Pmax = SP / self.Af + SM / Sf if Sf > 0 else SP / self.Af
            Pmin = SP / self.Af - SM / Sf if Sf > 0 else SP / self.Af
            res['Pmax_gross'] = Pmax
            res['Pmin_gross'] = Pmin
        else:
            # P(max) = L * SP / Af
            res['Pmax_gross'] = L * SP / self.Af
            res['Pmin_gross'] = 0
        
        # Net Pressure
        if isinstance(res.get('Pmax_gross'), str):
            res['Pmax_net'] = res['Pmax_gross']
        else:
            res['Pmax_net'] = res['Pmax_gross'] - (self.Ds + self.Tf) * self.gamma_s
        
        return res

    def _calc_bearing_flat(self, SP, SM, e_Df):
        """
        Bearing Pressure for Overturning about Axis through Flat Sides of Octagon.
        """
        res = {}
        
        # Section Modulus: Sf = If / (Df/2)
        Sf = self.If / (self.Df / 2) if self.Df > 0 else 0
        res['Sf'] = Sf
        
        # K interpolation from flat sides table
        threshold = 0.1321
        if e_Df > threshold:
            K, L = interpolate_K_L(e_Df, FLAT_SIDES_TABLE)
            if K is None:
                res['K'] = "Resize!"
                res['K_Df'] = "Resize!"
                res['pct_brg_area'] = "Resize!"
                res['L'] = "Resize!"
                res['Pmax_gross'] = "Resize!"
                res['Pmin_gross'] = "Resize!"
                res['Pmax_net'] = "Resize!"
                return res
        else:
            K = 0
            L = None
        
        res['K'] = K
        K_Df = K * self.Df if isinstance(K, (int, float)) else K
        res['K_Df'] = K_Df
        
        # %Brg Area calculation
        if isinstance(K_Df, (int, float)) and K_Df <= 0.62 * self.Df:
            pct_brg = self._calc_pct_brg_area_flat(K_Df)
        else:
            pct_brg = "Resize!"
        res['pct_brg_area'] = pct_brg
        
        # L coefficient
        if e_Df > threshold:
            res['L'] = L if L is not None else "Resize!"
        else:
            res['L'] = "N.A."
        
        # Gross Bearing Pressure
        if isinstance(res.get('L'), str) and res['L'] == "Resize!":
            res['Pmax_gross'] = "Resize!"
            res['Pmin_gross'] = "Resize!"
        elif e_Df <= threshold:
            Pmax = SP / self.Af + SM / Sf if Sf > 0 else SP / self.Af
            Pmin = SP / self.Af - SM / Sf if Sf > 0 else SP / self.Af
            res['Pmax_gross'] = Pmax
            res['Pmin_gross'] = Pmin
        else:
            res['Pmax_gross'] = L * SP / self.Af
            res['Pmin_gross'] = 0
        
        # Net Pressure
        if isinstance(res.get('Pmax_gross'), str):
            res['Pmax_net'] = res['Pmax_gross']
        else:
            res['Pmax_net'] = res['Pmax_gross'] - (self.Ds + self.Tf) * self.gamma_s
        
        return res

    def _calc_pct_brg_area_corners(self, K_Df):
        """
        Calculate %Bearing Area for axis through corners.
        Complex geometry based on where K*Df falls relative to octagon dimensions.
        """
        Cf = self.Cf
        Ef = self.Ef
        Af = self.Af
        
        # Region 1: K_Df <= Cf * sin(22.5°)
        limit1 = Cf * math.sin(math.radians(22.5))
        # Region 2: K_Df <= Ef/2
        limit2 = Ef / 2
        
        if K_Df <= limit1:
            # Simple triangle area subtracted
            unloaded_area = 0.5 * (K_Df / math.tan(math.radians(22.5)) * 2) * K_Df
            pct = (Af - unloaded_area) / Af * 100
        elif K_Df <= limit2:
            # More complex geometry
            term1 = (2 * Cf * math.cos(math.radians(22.5)) + 
                     (K_Df - Cf * math.sin(math.radians(22.5))) * math.tan(math.radians(22.5)))
            term2 = K_Df - Cf * math.sin(math.radians(22.5))
            triangle = 0.5 * (Cf * math.cos(math.radians(22.5)) * 2) * (Cf * math.sin(math.radians(22.5)))
            unloaded = term1 * term2 + triangle
            pct = (Af - unloaded) / Af * 100
        else:
            # Beyond Ef/2
            inside_width = (Ef - 2 * (K_Df - Ef / 2) * math.tan(math.radians(22.5)))
            term1 = (inside_width + (K_Df - Ef / 2) * math.tan(math.radians(22.5))) * (K_Df - Ef / 2)
            pct = (Af / 2 - term1) / Af * 100
        
        return pct

    def _calc_pct_brg_area_flat(self, K_Df):
        """
        Calculate %Bearing Area for axis through flat sides.
        """
        Cf = self.Cf
        Df = self.Df
        Af = self.Af
        
        # Region 1: K_Df <= Cf * cos(45°)
        limit1 = Cf * math.cos(math.radians(45))
        
        if K_Df <= limit1:
            unloaded = (Cf + K_Df * math.tan(math.radians(45))) * K_Df
            pct = (Af - unloaded) / Af * 100
        else:
            # K_Df > Cf*cos(45°)
            triangle = (Cf + Cf * math.sin(math.radians(45))) * (Cf * math.cos(math.radians(45)))
            rect = Df * (K_Df - Cf * math.cos(math.radians(45)))
            pct = (Af - (triangle + rect)) / Af * 100
        
        return pct

    def compute_ratios(self, P, H, M, Ds=None):
        """
        Compute analysis and return ratios for a single load combination.

        Args:
            P: Applied vertical load, kN
            H: Applied horizontal load, kN
            M: Applied moment, kN-m
            Ds: Soil depth override (if None, uses self.Ds)

        Returns:
            dict with 'FS_ot', 'FS_slid', 'Pmax_gross', 'Ratio_OT', 'Ratio_SLD', 'Ratio_SBC',
                  'Ratio_max', and full 'results'
        """
        # Override Ds if provided
        original_Ds = self.Ds
        if Ds is not None:
            self.Ds = Ds
            # Recompute weights with new Ds
            self.Ws = (self.Af - self.Ap) * self.Ds * self.gamma_s
        
        results = self.calculate(P, H, M)
        
        # FS factor: 1.5 if Ds < 1, else 1.0
        FS = 1.5 if self.Ds < 1 else 1.0
        
        # Compute ratios
        FS_ot = results['FS_ot']
        FS_slid = results['FS_slid']
        Pmax_gross = results['Pmax_gross_max']
        
        if isinstance(FS_ot, str):  # "N.A."
            Ratio_OT = 0.0
        else:
            Ratio_OT = FS / FS_ot if FS_ot != 0 else 0.0
        
        if isinstance(FS_slid, str):  # "N.A."
            Ratio_SLD = 0.0
        else:
            Ratio_SLD = FS / FS_slid if FS_slid != 0 else 0.0
        
        # Ratio_SBC = FS / Pmax_gross (matching Excel formula AG = Z/AD)
        if isinstance(Pmax_gross, str) or Pmax_gross == 0:
            Ratio_SBC = 0.0
        else:
            Ratio_SBC = FS / Pmax_gross
        
        Ratio_max = max(Ratio_OT, Ratio_SLD, Ratio_SBC)
        
        # Restore original Ds
        if Ds is not None:
            self.Ds = original_Ds
            self.Ws = (self.Af - self.Ap) * self.Ds * self.gamma_s
        
        return {
            'FS_ot': FS_ot,
            'FS_slid': FS_slid,
            'FS': FS,
            'Pmax_gross': results['Pmax_gross_max'],
            'Ratio_OT': Ratio_OT,
            'Ratio_SLD': Ratio_SLD,
            'Ratio_SBC': Ratio_SBC,
            'Ratio_max': Ratio_max,
            'results': results
        }
