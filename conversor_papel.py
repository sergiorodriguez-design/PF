575-    texto = rich_text(obj)
576-    runs = rich_runs(obj)
577-    if not texto and not runs:
578:        return f'    <w:p><w:pPr><w:pStyle w:val="{estilo}"/></w:pPr></w:p>'
579-    return (
580:        f'    <w:p><w:pPr><w:pStyle w:val="{estilo}"/></w:pPr>'
581-        f"{_runs_to_xml(runs, force_bold=negrita)}<\w:p>"
582-    )
583-
--
653-# Constructores XML
654-# =============================================================================
655-
656:def p(texto: str, estilo: str, negrita: bool = False) -> str:
657-    if isinstance(texto, dict):
658-        return p_rich(texto, estilo, negrita)
659-
--
662-    sp = ' xml:space="preserve"' if te and te != te.strip() else ""
663-
664-    if not te:
665:        return f'    <w:p><w:pPr><w:pStyle w:val="{estilo}"/></w:pPr></w:p>'
666-
667-    rpr = "<w:rPr><w:b/><w:bCs/></w:rPr>" if negrita else ""
668-
669-    return (
670:        f'    <w:p><w:pPr><w:pStyle w:val="{estilo}"/></w:pPr>'
671-        f"<w:r>{rpr}<w:t{sp}>{te}</w:t></w:r></w:p>"
672-    )
673-
--
680-        cleaned = _limpiar_vineta_rich(texto)
681-        runs = rich_runs(cleaned)
682-        return (
683:            f'    <w:p><w:pPr><w:pStyle w:val="{estilo}"/></w:pPr>'
684-            f"<w:r><w:t>{simbolo}</w:t></w:r>"
685-            f"<w:r><w:tab/></w:r>{_runs_to_xml(runs)}</w:p>"
686-        )
--
688-    te = esc(_limpiar_vineta_literal(texto))
689-
690-    return (
691:        f'    <w:p><w:pPr><w:pStyle w:val="{estilo}"/></w:pPr>'
692-        f"<w:r><w:t>{simbolo}</w:t></w:r>"
693-        f"<w:r><w:tab/><w:t>{te}</w:t></w:r></w:p>"
694-    )
--
699-        cleaned = _limpiar_vineta_rich(texto)
700-        runs = rich_runs(cleaned)
701-        return (
702:            f'    <w:p><w:pPr><w:pStyle w:val="Ejemplos-Vietanvl1"/></w:pPr>'
703-            f"<w:r><w:t>●</w:t></w:r>"
704-            f"<w:r><w:tab/></w:r>{_runs_to_xml(runs)}</w:p>"
705-        )
--
707-    te = esc(_limpiar_vineta_literal(texto))
708-
709-    return (
710:        f'    <w:p><w:pPr><w:pStyle w:val="Ejemplos-Vietanvl1"/></w:pPr>'
711-        f"<w:r><w:t>●</w:t></w:r>"
712-        f"<w:r><w:tab/><w:t>{te}</w:t></w:r></w:p>"
713-    )
--
718-        cleaned = _limpiar_vineta_rich(texto)
719-        runs = rich_runs(cleaned)
720-        return (
721:            f'    <w:p><w:pPr><w:pStyle w:val="Recuerda-Vietanvl1"/></w:pPr>'
722-            f"<w:r><w:t>●</w:t></w:r>"
723-            f"<w:r><w:tab/></w:r>{_runs_to_xml(runs)}</w:p>"
724-        )
725-    te = esc(_limpiar_vineta_literal(texto))
726-    return (
727:        f'    <w:p><w:pPr><w:pStyle w:val="Recuerda-Vietanvl1"/></w:pPr>'
728-        f"<w:r><w:t>●</w:t></w:r>"
729-        f"<w:r><w:tab/><w:t>{te}</w:t></w:r></w:p>"
730-    )
--
742-            nr["bold"] = True
743-            runs.append(nr)
744-        return (
745:            f'    <w:p><w:pPr><w:pStyle w:val="{estilo}"/></w:pPr>'
746-            f"<w:r><w:t>{simbolo}</w:t></w:r>"
747-            f"<w:r><w:tab/></w:r>{_runs_to_xml(runs)}</w:p>"
748-        )
749-
750-    te = esc(_limpiar_vineta_literal(texto))
751-    return (
752:        f'    <w:p><w:pPr><w:pStyle w:val="{estilo}"/></w:pPr>'
753-        f"<w:r><w:t>{simbolo}</w:t></w:r>"
754-        f"<w:r><w:tab/></w:r>"
755-        f"<w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t>{te}</w:t></w:r></w:p>"
--
773-        titulo_plain = titulo_str
774-        titulo_block = f'<w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t xml:space="preserve">{esc(titulo_str)}: </w:t></w:r>'
775-    header = (
776:        f'    <w:p><w:pPr><w:pStyle w:val="{estilo}"/></w:pPr>'
777-        f"<w:r><w:t>{simbolo}</w:t></w:r>"
778-        f"<w:r><w:tab/></w:r>"
779-        f"{titulo_block}"
--
820-        cleaned = _limpiar_vineta_rich(texto)
821-        runs = rich_runs(cleaned)
822-        return (
823:            '    <w:p><w:pPr><w:pStyle w:val="EjerciciosPregunta"/></w:pPr>'
824-            f'<w:r><w:t xml:space="preserve">{esc(letra)}) </w:t></w:r>'
825-            f"{_runs_to_xml(runs)}<\w:p>"
826-        )
827-    texto = re.sub(r"^[a-h][\).]\s*", "", str(texto).strip())
828-    return (
829:        '    <w:p><w:pPr><w:pStyle w:val="EjerciciosPregunta"/></w:pPr>'
830-        f"<w:r><w:t>{esc(letra)}) {esc(texto)}</w:t></w:r>"
831-        "<\w:p>"
832-    )
--
962-    if re.match(r"^Pie de imagen:\s*", texto, re.I):
963-        resto = re.sub(r"^(Pie de imagen:\s*)+", "", texto, flags=re.I)
964-        return (
965:            f'    <w:p><w:pPr><w:pStyle w:val="{style}"/></w:pPr>'
966-            f'<w:r><w:rPr><w:color w:val="{IMAGE_LABEL_RED}"/></w:rPr>'
967-            f'<w:t xml:space="preserve">Pie de imagen: </w:t></w:r>'
968-            f'<w:r><w:t>{esc(resto)}</w:t></w:r>'
--
973-        label = m.group(1)
974-        resto = re.sub(r"^(Descripci[oó]n de (la )?imagen:\s*)+", "", texto, flags=re.I)
975-        return (
976:            f'    <w:p><w:pPr><w:pStyle w:val="{style}"/></w:pPr>'
977-            f'<w:r><w:rPr><w:color w:val="{IMAGE_LABEL_RED}"/></w:rPr>'
978-            f'<w:t xml:space="preserve">{esc(label)} </w:t></w:r>'
979-            f'<w:r><w:t>{esc(resto)}</w:t></w:r>'
--
981-        )
982-    elif re.match(r"^Imagen_\d+$", texto, re.I):
983-        return (
984:            f'    <w:p><w:pPr><w:pStyle w:val="{style}"/></w:pPr>'
985-            f'<w:r><w:rPr><w:color w:val="{IMAGE_LABEL_RED}"/></w:rPr>'
986-            f'<w:t>{esc(texto)}</w:t></w:r>'
987-            f'<\w:p>'
--
3522-            raw = fila[i] if i < len(fila) else ""
3523-            paras = raw.split("\n") if raw else [""]
3524-            paras_xml = "".join(
3525:                f'<w:p><w:pPr><w:pStyle w:val="Cuerpoparrafo"/></w:pPr>'
3526-                f'<w:r><w:t xml:space="preserve">{esc(ln)}</w:t></w:r></w:p>'
3527-                for ln in paras
3528-            )
--
4068-            continue
4069-
4070-        # Strip source pStyle so graphics don't carry foreign styles into the output.
4071:        par_sin_estilo = re.sub(r'<w:pStyle w:val="[^"]+"/>', '', par)
4072-        graficos.append({
4073-            "xml": "    " + par_sin_estilo,
4074-            "prev": prev_txt,
