{-# LANGUAGE CPP #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE TemplateHaskell   #-}
{-# LANGUAGE RecordWildCards   #-}
{-# LANGUAGE DeriveGeneric #-}
module Codec.Xlsx.Types.PrintOptions where

#ifdef USE_MICROLENS
import Lens.Micro.TH (makeLenses)
#else
import Control.Lens (makeLenses)
#endif

import Control.DeepSeq (NFData)
import Data.Default
import qualified Data.Map as Map
import Data.Maybe (catMaybes)
import GHC.Generics (Generic)
import Text.XML

import Codec.Xlsx.Parser.Internal
import Codec.Xlsx.Writer.Internal

-- | See 18.3.1.70 "printOptions"
data PrintOptions = PrintOptions {
    _printOptionsHorizontalCentered :: Maybe Bool
  , _printOptionsVerticalCentered   :: Maybe Bool
  , _printOptionsHeadings           :: Maybe Bool
  , _printOptionsGridLines          :: Maybe Bool
  , _printOptionsGridLinesSet       :: Maybe Bool
  }
  deriving (Eq, Ord, Show, Generic)
instance NFData PrintOptions

instance Default PrintOptions where
  def = PrintOptions {
      _printOptionsHorizontalCentered = Nothing
    , _printOptionsVerticalCentered   = Nothing
    , _printOptionsHeadings           = Nothing
    , _printOptionsGridLines          = Nothing
    , _printOptionsGridLinesSet       = Nothing
    }

makeLenses ''PrintOptions

instance ToElement PrintOptions where
  toElement nm PrintOptions{..} = Element {
      elementName       = nm
    , elementNodes      = []
    , elementAttributes = Map.fromList . catMaybes $ [
          "horizontalCentered" .=? _printOptionsHorizontalCentered
        , "verticalCentered"   .=? _printOptionsVerticalCentered
        , "headings"           .=? _printOptionsHeadings
        , "gridLines"          .=? _printOptionsGridLines
        , "gridLinesSet"       .=? _printOptionsGridLinesSet
        ]
    }

instance FromCursor PrintOptions where
    fromCursor cur = do
      _printOptionsHorizontalCentered <- maybeAttribute "horizontalCentered" cur
      _printOptionsVerticalCentered   <- maybeAttribute "verticalCentered" cur
      _printOptionsHeadings           <- maybeAttribute "headings" cur
      _printOptionsGridLines          <- maybeAttribute "gridLines" cur
      _printOptionsGridLinesSet       <- maybeAttribute "gridLinesSet" cur
      return PrintOptions{..}

instance FromXenoNode PrintOptions where
    fromXenoNode root =
      parseAttributes root $ do
        _printOptionsHorizontalCentered <- maybeAttr "horizontalCentered"
        _printOptionsVerticalCentered   <- maybeAttr "verticalCentered"
        _printOptionsHeadings           <- maybeAttr "headings"
        _printOptionsGridLines          <- maybeAttr "gridLines"
        _printOptionsGridLinesSet       <- maybeAttr "gridLinesSet"
        return PrintOptions{..}
