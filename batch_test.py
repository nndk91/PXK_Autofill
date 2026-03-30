#!/usr/bin/env python3
"""Batch evaluation for App v4 against the labeled training folders."""

import os

from pxk_core_v4 import (
    extract_folder_pxk_data,
    find_forms_in_folder,
    load_reference_scorer,
    match_pxk_v4,
    read_filled_pxk_values,
    read_form_rows_from_file,
)


def test_folder(folder_path):
    folder_name = os.path.basename(folder_path)
    print(f"\n{'='*60}")
    print(f"Testing folder: {folder_name}")
    print('='*60)

    empty_form, filled_form = find_forms_in_folder(folder_path)
    if not empty_form or not filled_form:
        print(f"❌ Missing forms in {folder_name}")
        return None

    print(f"Empty form: {os.path.basename(empty_form)}")
    print(f"Filled form: {os.path.basename(filled_form)}")

    pxk_folder = os.path.join(folder_path, 'PXK')
    pdf_files = [f for f in os.listdir(pxk_folder) if f.endswith('.pdf')]
    print(f"PDF files: {len(pdf_files)}")
    pxk_totals, _, pxk_do_no, errors = extract_folder_pxk_data(folder_path)

    print(f"Extracted PXKs: {len(pxk_totals)}")
    if errors:
        print(f"Errors: {len(errors)}")

    form_rows = read_form_rows_from_file(empty_form)
    print(f"Form rows: {len(form_rows)}")
    filled_values = read_filled_pxk_values(filled_form)

    print(f"Ground truth rows: {len(filled_values)}")

    scorer = load_reference_scorer(os.path.dirname(folder_path))
    result, status, note_pxks = match_pxk_v4(form_rows, pxk_totals, pxk_do_no, scorer=scorer)

    correct = 0
    wrong = 0
    missing = 0
    wrong_details = []

    for fr in form_rows:
        i = fr['idx']
        row = fr['row']
        matched_pxk = result[i]
        expected = filled_values.get(row)

        if expected:
            if matched_pxk:
                exp_normalized = expected.lstrip('0').replace('.0', '')
                match_normalized = str(matched_pxk).lstrip('0').replace('.0', '')
                if exp_normalized == match_normalized:
                    correct += 1
                else:
                    wrong += 1
                    wrong_details.append({
                        'row': row,
                        'mh': fr['ma_hang'],
                        'sl': fr['sl'],
                        'inv': fr['inv'],
                        'expected': expected,
                        'got': matched_pxk,
                        'status': status[i]
                    })
            else:
                missing += 1

    total_with_expected = correct + wrong + missing

    metrics = {
        'folder': folder_name,
        'total_rows': len(form_rows),
        'expected_rows': total_with_expected,
        'correct': correct,
        'wrong': wrong,
        'missing': missing,
        'auto_matched': sum(1 for s in status if s == 'auto'),
        'ambiguous': sum(1 for s in status if s == 'ambiguous'),
        'no_match': sum(1 for s in status if s == 'no_match'),
        'accuracy': correct / total_with_expected * 100 if total_with_expected > 0 else 0,
        'match_rate': (correct + wrong) / total_with_expected * 100 if total_with_expected > 0 else 0,
        'wrong_details': wrong_details[:10],
    }

    print(f"\nResults:")
    print(f"  Auto-matched: {metrics['auto_matched']}")
    print(f"  Ambiguous: {metrics['ambiguous']}")
    print(f"  No match: {metrics['no_match']}")
    print(f"  Correct: {metrics['correct']}")
    print(f"  Wrong: {metrics['wrong']}")
    print(f"  Missing: {metrics['missing']}")
    print(f"  Accuracy: {metrics['accuracy']:.1f}%")
    print(f"  Match rate: {metrics['match_rate']:.1f}%")

    if wrong_details:
        print(f"\nSample wrong matches:")
        for w in wrong_details[:5]:
            print(f"  Row {w['row']}: {w['mh']} (SL={w['sl']}, Inv={w['inv']})")
            print(f"    Expected: {w['expected']}, Got: {w['got']} [{w['status']}]")

    return metrics


if __name__ == '__main__':
    folders = [
        '1544-1584',
        '1585-1640',
        '1641-1700',
        '1742-1790'
    ]

    all_metrics = []

    for folder in folders:
        folder_path = os.path.join('/Users/nndk91/Data/SMC/XNK_PXK_Program', folder)
        if os.path.exists(folder_path):
            metrics = test_folder(folder_path)
            if metrics:
                all_metrics.append(metrics)
        else:
            print(f"❌ Folder not found: {folder}")

    # Summary
    print("\n" + "="*60)
    print("SUMMARY ACROSS ALL FOLDERS")
    print("="*60)

    if all_metrics:
        total_correct = sum(m['correct'] for m in all_metrics)
        total_wrong = sum(m['wrong'] for m in all_metrics)
        total_missing = sum(m['missing'] for m in all_metrics)
        total_expected = sum(m['expected_rows'] for m in all_metrics)

        print(f"\nTotal folders: {len(all_metrics)}")
        print(f"Total expected rows: {total_expected}")
        print(f"Total correct: {total_correct}")
        print(f"Total wrong: {total_wrong}")
        print(f"Total missing: {total_missing}")
        print(f"Overall accuracy: {total_correct/total_expected*100:.1f}%")
        print(f"Overall match rate: {(total_correct+total_wrong)/total_expected*100:.1f}%")

        print("\nPer-folder breakdown:")
        for m in all_metrics:
            print(f"  {m['folder']}: {m['accuracy']:.1f}% accuracy, {m['match_rate']:.1f}% match rate")
