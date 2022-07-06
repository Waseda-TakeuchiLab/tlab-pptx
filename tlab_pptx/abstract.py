# Copyright (c) 2022 Shuhei Nitta. All rights reserved.
import abc

import pptx.presentation

from tlab_pptx import common


class AbstractPresentation(abc.ABC):
    """Abstract class for Presentation."""

    @abc.abstractmethod
    def build(self) -> pptx.presentation.Presentation:
        """Build a Presentation object

        Returns
        -------
        pptx.presentaion.Presentation
            A Presentation object of python-pptx
        """

    @abc.abstractmethod
    def save(self, filepath_or_buffer: common.FilePathOrBuffer) -> None:
        """Save as a `pptx` file.

        Parameters
        ----------
        filepath_or_buffer : tlab_pptx.typing.FilePathOrBuffer
            A filepath string or buffer object.
        """
