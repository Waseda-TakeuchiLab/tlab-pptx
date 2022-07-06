# Copyright (c) 2022 Shuhei Nitta. All rights reserved.
from unittest import TestCase, mock
import doctest
import datetime
import tempfile
import pathlib

import plotly.graph_objects as go

from tlab_pptx import (
    photo_luminescence as pl,
    common
)


class TestPresentation_build(TestCase):  # TODO: Implement unittests
    pass


class TestPresentation_save(TestCase):

    def _test(
        self,
        filepath_or_buffer: common.FilePathOrBuffer = "path/to/example.pptx"
    ) -> None:
        prs = pl.Presentation(
            title="title",
            excitation_wavelength=400,
            excitation_power=1,
            time_range=10,
            center_wavelength=480,
            FWHM=48,
            frame=10000,
            date=datetime.date(2022, 1, 1),
            h_fig=mock.Mock(spec_set=go.Figure),
            v_fig=mock.Mock(spec_set=go.Figure),
            a=60,
            b=40,
            tau1=1.0,
            tau2=3.0
        )
        with mock.patch("tlab_pptx.photo_luminescence.Presentation.build") as build_mock:
            prs.save(filepath_or_buffer)
            build_mock.return_value.save.assert_called_once_with(
                filepath_or_buffer
            )

    def test_filepath_or_buffer(self) -> None:
        filepaths: list[common.FilePath] = [
            "test_presentation_save.pptx",
            pathlib.Path("test_presentation_save.pptx"),
        ]
        for filepath in filepaths:
            with self.subTest(filepath=filepath):
                self._test(filepath_or_buffer=filepath)
        with tempfile.TemporaryDirectory() as tmpdir:
            _tmpdir = pathlib.Path(tmpdir)
            for filepath in filepaths:
                with open(_tmpdir / filepath, "wb") as f:
                    self._test(filepath_or_buffer=f)


def load_tests(loader, tests, _):  # type: ignore
    tests.addTests(doctest.DocTestSuite(pl))
    return tests
