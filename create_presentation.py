#!/usr/bin/env python3
"""
Script pour créer une présentation PowerPoint sur les stratégies de développement économique
Nécessite: pip install python-pptx
"""

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.enum.text import PP_ALIGN
    from pptx.dml.color import RGBColor
except ImportError:
    print("ERREUR: La bibliothèque python-pptx n'est pas installée.")
    print("Veuillez l'installer avec: pip install python-pptx")
    exit(1)

def add_title_slide(prs, title, subtitle):
    """Ajoute une diapositive de titre"""
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    
    title_shape = slide.shapes.title
    subtitle_shape = slide.placeholders[1]
    
    title_shape.text = title
    subtitle_shape.text = subtitle
    
    return slide

def add_content_slide(prs, title, content_points):
    """Ajoute une diapositive avec titre et points"""
    slide_layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(slide_layout)
    
    title_shape = slide.shapes.title
    title_shape.text = title
    
    body_shape = slide.placeholders[1]
    text_frame = body_shape.text_frame
    text_frame.clear()
    
    for point in content_points:
        p = text_frame.add_paragraph()
        p.text = point
        p.level = 0
        p.space_before = Pt(6)
        
    return slide

def create_presentation():
    """Crée la présentation complète"""
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    
    # Diapositive 1: Page de titre
    add_title_slide(
        prs,
        "Stratégies de Développement Économique",
        "Approches Agricoles et Manufacturières"
    )
    
    # Diapositive 2: Introduction
    add_content_slide(
        prs,
        "Introduction",
        [
            "Le développement économique est un processus complexe et multidimensionnel",
            "Deux stratégies principales dominent historiquement:",
            "  • Stratégie agricole: base traditionnelle de l'économie",
            "  • Stratégie manufacturière: moteur de l'industrialisation",
            "Ces approches ont façonné la croissance des nations modernes",
            "Chaque pays adapte sa stratégie selon ses ressources et contexte"
        ]
    )
    
    # Diapositive 3: Contexte historique
    add_content_slide(
        prs,
        "Contexte Historique",
        [
            "Révolution agricole: première transformation économique majeure",
            "Révolution industrielle: passage vers la manufacture (18e-19e siècles)",
            "Exemples historiques:",
            "  • Royaume-Uni: pionnier de l'industrialisation",
            "  • États-Unis: combinaison agriculture-industrie",
            "  • Japon: industrialisation rapide post-Meiji",
            "Les pays en développement cherchent leur propre voie"
        ]
    )
    
    # Diapositive 4: Stratégie Agricole - Définition
    add_content_slide(
        prs,
        "Stratégie Agricole: Fondements",
        [
            "Définition: Développement basé sur le secteur primaire",
            "Objectifs principaux:",
            "  • Sécurité alimentaire nationale",
            "  • Création d'emplois ruraux",
            "  • Génération de revenus d'exportation",
            "  • Base pour l'industrialisation future",
            "Secteurs clés: cultures vivrières, cultures de rente, élevage, pêche"
        ]
    )
    
    # Diapositive 5: Stratégie Agricole - Avantages
    add_content_slide(
        prs,
        "Avantages de la Stratégie Agricole",
        [
            "✓ Utilisation des ressources naturelles disponibles",
            "✓ Faible investissement initial en capital",
            "✓ Absorption de la main-d'œuvre rurale",
            "✓ Réduction de la pauvreté en milieu rural",
            "✓ Développement équilibré du territoire",
            "✓ Préservation des traditions et cultures locales",
            "✓ Durabilité environnementale (si bien gérée)"
        ]
    )
    
    # Diapositive 6: Stratégie Agricole - Défis
    add_content_slide(
        prs,
        "Défis de la Stratégie Agricole",
        [
            "✗ Productivité limitée sans modernisation",
            "✗ Vulnérabilité aux conditions climatiques",
            "✗ Fluctuations des prix internationaux",
            "✗ Accès limité aux technologies modernes",
            "✗ Infrastructures rurales insuffisantes",
            "✗ Difficulté à générer une forte valeur ajoutée",
            "✗ Exode rural vers les villes"
        ]
    )
    
    # Diapositive 7: Exemples de pays - Agriculture
    add_content_slide(
        prs,
        "Exemples: Stratégie Agricole",
        [
            "Vietnam: Réformes agricoles (Doi Moi) → 2e exportateur de riz",
            "Côte d'Ivoire: Leader mondial du cacao",
            "Nouvelle-Zélande: Agriculture intensive et exportations",
            "Brésil: Agrobusiness moderne et compétitif",
            "Thaïlande: Diversification agricole réussie",
            "Éthiopie: Investissements massifs dans l'agriculture"
        ]
    )
    
    # Diapositive 8: Stratégie Manufacturière - Définition
    add_content_slide(
        prs,
        "Stratégie Manufacturière: Fondements",
        [
            "Définition: Développement par l'industrialisation",
            "Objectifs principaux:",
            "  • Transformation des matières premières",
            "  • Création de valeur ajoutée élevée",
            "  • Développement technologique",
            "  • Urbanisation et modernisation",
            "Secteurs: textile, électronique, automobile, chimie, métallurgie"
        ]
    )
    
    # Diapositive 9: Stratégie Manufacturière - Avantages
    add_content_slide(
        prs,
        "Avantages de la Stratégie Manufacturière",
        [
            "✓ Forte valeur ajoutée et productivité élevée",
            "✓ Création d'emplois qualifiés et mieux rémunérés",
            "✓ Transfert et développement technologique",
            "✓ Diversification économique",
            "✓ Augmentation des exportations à haute valeur",
            "✓ Effet d'entraînement sur d'autres secteurs",
            "✓ Croissance économique rapide"
        ]
    )
    
    # Diapositive 10: Stratégie Manufacturière - Défis
    add_content_slide(
        prs,
        "Défis de la Stratégie Manufacturière",
        [
            "✗ Investissements initiaux massifs requis",
            "✗ Besoin de main-d'œuvre qualifiée",
            "✗ Dépendance aux technologies étrangères",
            "✗ Concurrence internationale intense",
            "✗ Impact environnemental significatif",
            "✗ Risque de déséquilibres régionaux",
            "✗ Vulnérabilité aux crises économiques mondiales"
        ]
    )
    
    # Diapositive 11: Exemples de pays - Manufacture
    add_content_slide(
        prs,
        "Exemples: Stratégie Manufacturière",
        [
            "Chine: 'Usine du monde' - croissance spectaculaire",
            "Corée du Sud: Des textiles à l'électronique de pointe",
            "Taïwan: Semi-conducteurs et haute technologie",
            "Allemagne: Excellence en ingénierie et automobile",
            "Bangladesh: Industrie textile en expansion",
            "Mexique: Manufacture pour le marché nord-américain"
        ]
    )
    
    # Diapositive 12: Approches Hybrides
    add_content_slide(
        prs,
        "Stratégies Hybrides et Modernes",
        [
            "Tendance actuelle: Combinaison des deux approches",
            "Agro-industrie: Transformation des produits agricoles",
            "Industrialisation rurale: Usines en zones agricoles",
            "Chaînes de valeur intégrées: De la ferme à l'usine",
            "Exemples réussis:",
            "  • Inde: IT + agriculture",
            "  • Malaisie: Huile de palme + électronique",
            "  • Chili: Mines + agro-exportations"
        ]
    )
    
    # Diapositive 13: Facteurs de Succès
    add_content_slide(
        prs,
        "Facteurs Clés de Succès",
        [
            "Stabilité politique et bonne gouvernance",
            "Investissement dans l'éducation et la formation",
            "Infrastructures de qualité (routes, ports, énergie)",
            "Politiques économiques cohérentes et prévisibles",
            "Ouverture au commerce international",
            "Innovation et recherche-développement",
            "Institutions fortes et État de droit"
        ]
    )
    
    # Diapositive 14: Rôle de l'État
    add_content_slide(
        prs,
        "Rôle de l'État dans le Développement",
        [
            "Planification stratégique à long terme",
            "Investissements dans les infrastructures",
            "Politiques de soutien ciblées (subventions, crédits)",
            "Régulation et création d'un environnement favorable",
            "Protection sociale et redistribution",
            "Promotion de l'innovation et de la R&D",
            "Équilibre entre intervention et libre marché"
        ]
    )
    
    # Diapositive 15: Défis Contemporains
    add_content_slide(
        prs,
        "Défis du 21e Siècle",
        [
            "Changement climatique et durabilité environnementale",
            "Révolution numérique et automatisation",
            "Mondialisation et concurrence accrue",
            "Inégalités croissantes (urbain/rural, riches/pauvres)",
            "Transition énergétique nécessaire",
            "Sécurité alimentaire face à la croissance démographique",
            "Adaptation aux nouvelles technologies (IA, robotique)"
        ]
    )
    
    # Diapositive 16: Perspectives d'Avenir
    add_content_slide(
        prs,
        "Perspectives et Tendances Futures",
        [
            "Économie verte: Agriculture durable + industrie propre",
            "Digitalisation: Agriculture de précision, Industrie 4.0",
            "Économie circulaire: Réduction des déchets, recyclage",
            "Développement inclusif: Réduction des inégalités",
            "Intégration régionale: Marchés communs, coopération",
            "Innovation sociale: Entrepreneuriat, coopératives",
            "Résilience: Capacité d'adaptation aux chocs"
        ]
    )
    
    # Diapositive 17: Conclusion - Partie 1
    add_content_slide(
        prs,
        "Conclusion: Synthèse",
        [
            "Aucune stratégie unique ne convient à tous les pays",
            "Le contexte national détermine l'approche optimale:",
            "  • Ressources disponibles",
            "  • Niveau de développement initial",
            "  • Position géographique",
            "  • Capital humain",
            "Les stratégies évoluent avec le temps et le développement"
        ]
    )
    
    # Diapositive 18: Conclusion - Partie 2
    add_content_slide(
        prs,
        "Conclusion: Leçons Apprises",
        [
            "L'agriculture reste fondamentale pour la sécurité alimentaire",
            "L'industrialisation accélère la croissance économique",
            "Les approches hybrides offrent le meilleur équilibre",
            "L'investissement humain est crucial dans tous les cas",
            "La durabilité doit être intégrée dès le départ",
            "La coopération internationale facilite le développement",
            "L'adaptation et l'innovation sont essentielles au succès"
        ]
    )
    
    # Diapositive 19: Recommandations
    add_content_slide(
        prs,
        "Recommandations Stratégiques",
        [
            "Pour les pays en développement:",
            "  • Commencer par moderniser l'agriculture",
            "  • Investir massivement dans l'éducation",
            "  • Développer progressivement l'industrie",
            "  • Attirer les investissements étrangers stratégiques",
            "  • Créer des zones économiques spéciales",
            "  • Favoriser l'entrepreneuriat local",
            "  • Intégrer les objectifs de développement durable"
        ]
    )
    
    # Diapositive 20: Fin
    add_title_slide(
        prs,
        "Merci de votre attention",
        "Questions et Discussion"
    )
    
    # Sauvegarder la présentation
    filename = "Strategies_Developpement_Economique.pptx"
    prs.save(filename)
    print(f"✓ Présentation créée avec succès: {filename}")
    print(f"✓ Nombre de diapositives: {len(prs.slides)}")
    return filename

if __name__ == "__main__":
    try:
        create_presentation()
    except Exception as e:
        print(f"ERREUR lors de la création de la présentation: {e}")
        exit(1)
