"""
better_abc adds decorator abstract_attribute to the abc module.

Thanks to krassowski (stackoverflow) who provided this code as CC0
"""
from abc import ABCMeta as NativeABCMeta


def abstract_attribute(obj=None):
    if obj is None:
        obj = object()
    obj.__is_abstract_attribute__ = True
    return obj


class ABCMeta(NativeABCMeta):
    def __call__(cls, *args, **kwargs):
        instance = NativeABCMeta.__call__(cls, *args, **kwargs)
        abstract_attributes = {
            name
            for name in dir(instance)
            if getattr(getattr(instance, name), '__is_abstract_attribute__', False)
        }
        if abstract_attributes:
            raise NotImplementedError(
                "Can't instantiate abstract class {} with"
                " abstract attributes: {}".format(
                    cls.__name__,
                    ', '.join(abstract_attributes)
                )
            )
        return instance
