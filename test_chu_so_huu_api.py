import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import Mock

from openpyxl import Workbook

from chu_so_huu_api import (
    TIMEOUT,
    URL_CAP_NHAT_PHAN_LOAI_THUA_DAT,
    URL_TRA_CUU_THUA_DAT,
    build_cap_nhat_loai_doi_tuong_payload,
    build_cap_nhat_payloads_from_response,
    build_tra_cuu_payload,
    cap_nhat_loai_doi_tuong,
    doc_danh_sach_thua,
    lay_tinh_hinh_dang_ky_ids_thieu_loai_doi_tuong,
    tra_cuu_va_cap_nhat_loai_doi_tuong,
    tra_cuu_thua_dat,
)


class BuildTraCuuPayloadTests(unittest.TestCase):
    def test_builds_expected_payload(self):
        payload = build_tra_cuu_payload(
            so_thua=118,
            so_to=187,
            xa_id=24373,
            tinh_id=66,
        )

        self.assertEqual(
            payload,
            {
                "traCuu[soThuTuThua]": "118",
                "traCuu[soHieuToBanDo]": "187",
                "traCuu[soPhatHanh]": "",
                "traCuu[hoTenChu]": "",
                "traCuu[phanLoai]": "-1",
                "traCuu[type]": "-1",
                "traCuu[loaiChu]": "-1",
                "traCuu[tuNgay]": "",
                "traCuu[denNgay]": "",
                "traCuu[query]": "",
                "traCuu[xaId]": "24373",
                "traCuu[huyenId]": "0",
                "traCuu[tinhId]": "66",
                "start": "0",
                "length": "10",
                "exportWard": "false",
                "subLength": "3000",
                "sort[Field]": "_id",
                "sort[Direction]": "1",
            },
        )

    def test_posts_form_payload_to_search_endpoint(self):
        response = Mock()
        response.json.return_value = {"success": True, "data": []}
        session = Mock()
        session.post.return_value = response

        result = tra_cuu_thua_dat(
            session,
            so_thua=118,
            so_to=187,
            xa_id=24373,
            tinh_id=66,
        )

        self.assertEqual(result, {"success": True, "data": []})
        response.raise_for_status.assert_called_once_with()
        session.post.assert_called_once()
        _, kwargs = session.post.call_args
        self.assertEqual(session.post.call_args.args[0], URL_TRA_CUU_THUA_DAT)
        self.assertEqual(kwargs["data"]["traCuu[soThuTuThua]"], "118")
        self.assertEqual(kwargs["data"]["traCuu[soHieuToBanDo]"], "187")
        self.assertEqual(
            kwargs["headers"]["Content-Type"],
            "application/x-www-form-urlencoded; charset=UTF-8",
        )
        self.assertEqual(kwargs["timeout"], TIMEOUT)

    def test_selects_registration_id_missing_owner_object_type(self):
        response_data = {
            "success": True,
            "data": [
                {
                    "tinhHinhDangKyId": 20733604,
                    "thongTinDangKyChuaDapUngNhom1": [
                        "TINHHINHDANGKY.20733604|CANHAN.300038_10|loaiDoiTuongId"
                    ],
                }
            ],
            "recordsTotal": 1,
        }

        self.assertEqual(
            lay_tinh_hinh_dang_ky_ids_thieu_loai_doi_tuong(response_data),
            [20733604],
        )

    def test_ignores_other_errors_and_mismatched_registration_ids(self):
        response_data = {
            "success": True,
            "data": [
                {
                    "tinhHinhDangKyId": 100,
                    "thongTinDangKyChuaDapUngNhom1": [
                        "TINHHINHDANGKY.100|CANHAN.1_0|soGiayTo"
                    ],
                },
                {
                    "tinhHinhDangKyId": 200,
                    "thongTinDangKyChuaDapUngNhom1": [
                        "TINHHINHDANGKY.201|CANHAN.2_0|loaiDoiTuongId"
                    ],
                },
            ],
        }

        self.assertEqual(
            lay_tinh_hinh_dang_ky_ids_thieu_loai_doi_tuong(response_data),
            [],
        )

    def test_builds_owner_object_type_update_payload(self):
        row = {
            "id": "6a61186054b88b7df52e4947",
            "tinhHinhDangKyId": 20733604,
            "thongTinDangKyChuaDapUngNhom1": [
                "TINHHINHDANGKY.20733604|CANHAN.300038_10|loaiDoiTuongId"
            ],
        }

        payload = build_cap_nhat_loai_doi_tuong_payload(row)

        self.assertEqual(
            payload,
            {
                "id": "6a61186054b88b7df52e4947",
                "tinhHinhDangKyId": 20733604,
                "data": (
                    '{"TINHHINHDANGKY.20733604|CANHAN.300038_10|'
                    'loaiDoiTuongId":"22"}'
                ),
            },
        )

    def test_builds_one_update_for_both_people_in_married_couple(self):
        row = {
            "id": "6a61189d54b88b7df52eb980",
            "tinhHinhDangKyId": 20733804,
            "thongTinDangKyChuaDapUngNhom1": [
                (
                    "TINHHINHDANGKY.20733804|VOCHONG.4218522_0|"
                    "CANHAN.306077_10|loaiDoiTuongId"
                ),
                (
                    "TINHHINHDANGKY.20733804|VOCHONG.4218522_0|"
                    "CANHAN.306066_10|loaiDoiTuongId"
                ),
            ],
        }

        payload = build_cap_nhat_loai_doi_tuong_payload(row)

        self.assertEqual(payload["id"], "6a61189d54b88b7df52eb980")
        self.assertEqual(payload["tinhHinhDangKyId"], 20733804)
        self.assertEqual(
            payload["data"],
            (
                '{"TINHHINHDANGKY.20733804|VOCHONG.4218522_0|'
                'CANHAN.306077_10|loaiDoiTuongId":"22",'
                '"TINHHINHDANGKY.20733804|VOCHONG.4218522_0|'
                'CANHAN.306066_10|loaiDoiTuongId":"22"}'
            ),
        )

    def test_skips_rows_without_missing_owner_object_type(self):
        response_data = {
            "success": True,
            "data": [
                {
                    "id": "record-1",
                    "tinhHinhDangKyId": 100,
                    "thongTinDangKyChuaDapUngNhom1": [
                        "TINHHINHDANGKY.100|CANHAN.1_0|soGiayTo"
                    ],
                }
            ],
        }

        self.assertEqual(build_cap_nhat_payloads_from_response(response_data), [])

    def test_posts_update_payload_as_json(self):
        response = Mock()
        response.json.return_value = {"success": True}
        session = Mock()
        session.post.return_value = response
        payload = {
            "id": "6a61186054b88b7df52e4947",
            "tinhHinhDangKyId": 20733604,
            "data": (
                '{"TINHHINHDANGKY.20733604|CANHAN.300038_10|'
                'loaiDoiTuongId":"22"}'
            ),
        }

        result = cap_nhat_loai_doi_tuong(session, payload)

        self.assertEqual(result, {"success": True})
        response.raise_for_status.assert_called_once_with()
        session.post.assert_called_once_with(
            URL_CAP_NHAT_PHAN_LOAI_THUA_DAT,
            json=payload,
            headers={"Content-Type": "application/json; charset=UTF-8"},
            timeout=TIMEOUT,
        )

    def test_full_flow_searches_then_updates_without_intermediate_request(self):
        search_response = Mock()
        search_response.json.return_value = {
            "success": True,
            "data": [
                {
                    "id": "6a61186054b88b7df52e4947",
                    "tinhHinhDangKyId": 20733604,
                    "thongTinDangKyChuaDapUngNhom1": [
                        "TINHHINHDANGKY.20733604|CANHAN.300038_10|loaiDoiTuongId"
                    ],
                }
            ],
        }
        update_response = Mock()
        update_response.json.return_value = {"success": True}
        session = Mock()
        session.post.side_effect = [search_response, update_response]

        results = tra_cuu_va_cap_nhat_loai_doi_tuong(
            session,
            so_thua=118,
            so_to=187,
            xa_id=24373,
            tinh_id=66,
        )

        self.assertEqual(results, [{"success": True}])
        self.assertEqual(session.post.call_count, 2)
        self.assertEqual(session.post.call_args_list[0].args[0], URL_TRA_CUU_THUA_DAT)
        self.assertEqual(
            session.post.call_args_list[1].args[0],
            URL_CAP_NHAT_PHAN_LOAI_THUA_DAT,
        )

    def test_reads_vietnamese_excel_headers_and_removes_duplicates(self):
        with TemporaryDirectory() as temp_dir:
            input_path = Path(temp_dir) / "input.xlsx"
            workbook = Workbook()
            worksheet = workbook.active
            worksheet.append(["STT", "Số tờ", "Số thửa"])
            worksheet.append([1, 187, 118])
            worksheet.append([2, 187, 118])
            worksheet.append([3, 188, 119])
            workbook.save(input_path)

            parcels = doc_danh_sach_thua(str(input_path))

        self.assertEqual(
            parcels,
            [
                {"excel_row": 2, "so_to": "187", "so_thua": "118"},
                {"excel_row": 4, "so_to": "188", "so_thua": "119"},
            ],
        )


if __name__ == "__main__":
    unittest.main()
